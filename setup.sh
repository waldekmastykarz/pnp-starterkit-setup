#!/usr/bin/env bash

set -o errexit
set -o pipefail
set -o nounset
# set -o xtrace

# Set magic variables for current file & dir
__dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
__file="${__dir}/$(basename "${BASH_SOURCE[0]}")"
__base="$(basename ${__file} .sh)"
__root="$(cd "$(dirname "${__dir}")" && pwd)"

# helper functions
. ./functions.sh

# script arguments
arg1="${1:-}"

tenantUrl="https://m365x526922.sharepoint.com"
prefix=""
skipSolutionDeployment=false
skipSiteCreation=false

# create sites

portalUrl=$tenantUrl/sites/$(echo $prefix)portal
echo $portalUrl
exit 1
if [ ! $skipSiteCreation ]; then
  echo "Provisioning portal site at $portalUrl..."
  site=$(o365 spo site get --url $portalUrl --output json)
  if $(isError "$site"); then
    echo "Creating site..."
    o365 spo site add --type CommunicationSite --url $portalUrl --title 'PnP SP Starter Kit' --description 'PnP SP Starket Kit Hub'
    success "DONE"
  else
    warning "Site already exists"
  fi

  hubsiteId=$(o365 spo hubsite list --output json | jq -r '.[] | select(.SiteUrl == "'"$portalUrl"'") | .ID')
  if [ -z "$hubsiteId" ]; then
    echo "Hubsite not found. Registering..."
    out=$(o365 spo hubsite register --url $portalUrl --output json)
    if $(isError "$out"); then
      errorMessage "$out"
      exit 1
    fi
    hubsiteId=$(echo "$out" | jq -r '.ID')
    success "DONE"
  else
    warning "Hubsite already exists"
  fi

  siteUrl=$tenantUrl/sites/$(echo $prefix)hr
  echo "Provisioning site at $siteUrl..."
  site=$(o365 spo site get --url $siteUrl --output json)
  if $(isError "$site"); then
    echo "Creating site..."
    o365 spo site add --type TeamSite --url $siteUrl --alias $(echo $prefix)hr --title 'Human Resources' --description 'Human Resources'
    success "DONE"
  else
    warning "Site already exists"
  fi
  echo "Connecting site to the hub site..."
  o365 spo hubsite connect --url $siteUrl --hubSiteId $hubsiteId
  success "DONE"

  siteUrl=$tenantUrl/sites/$(echo $prefix)marketing
  echo "Provisioning site at $siteUrl..."
  site=$(o365 spo site get --url $siteUrl --output json)
  if $(isError "$site"); then
    echo "Creating site..."
    o365 spo site add --type TeamSite --url $siteUrl --alias $(echo $prefix)marketing --title 'Marketing' --description 'Marketing'
    success "DONE"
  else
    warning "Site already exists"
  fi
  echo "Connecting site to the hub site..."
  o365 spo hubsite connect --url $siteUrl --hubSiteId $hubsiteId
  success "DONE"
fi

if [ ! $skipSolutionDeployment ]; then
  echo "Deploying solution..."

  app=$(o365 spo app get --name "sharepoint-portal-showcase.sppkg" --output json)
  if [ ! -z "$app" ]; then
    echo "Solution package already exists. Removing..."
    o365 spo app remove --name "sharepoint-portal-showcase.sppkg"
    success "DONE"
  fi

  echo "Add solution package to tenant app catalog..."
  o365 spo app add --filePath ./sharepoint-portal-showcase.sppkg
  success "DONE"

  echo "Enable Office 365 Public CDN..."
  o365 spo cdn set --enabled true --type Public
  success "DONE"

  echo "Configure */CLIENTSIDEASSETS Public CDN origin..."
  origin=$(o365 spo cdn origin list --output json | jq '.[] | select(. == "*/CLIENTSIDEASSETS")')
  if [ -z "$origin" ]; then
    echo "Origin doesn't exist. Configuring..."
    o365 spo cdn origin add --origin '*/CLIENTSIDEASSETS' --type 'Public'
    success "DONE"
  else
    success "Origin already exists"
  fi

  echo "Configure web API permissions..."
  echo "Retrieving SharePoint Online Client Extensibility service principal..."
  sp=$(o365 aad sp get --displayName 'SharePoint Online Client Extensibility' --output json | jq -r '.objectId')
  if [ -z "$sp" ]; then
    error "SharePoint Online Client Extensibility service principal not found in Azure AD"
    exit 1
  fi
  success "DONE"
  echo "Retrieving Microsoft Graph service principal..."
  # need to store in a temp file because the output is too long to fit directly
  # into a variable
  o365 aad sp get --displayName 'Microsoft Graph' --output json > .tmp
  graph=$(cat .tmp | jq -r '.objectId')
  # clean up temp file
  rm .tmp
  if [ -z "$graph" ]; then
    error "Microsoft Graph service principal not found in Azure AD"
    exit 1
  fi
  success "DONE"
  echo "Retrieving SharePoint Online Client Extensibility service principal permissions for MS Graph..."
  graphPermissions=$(o365 aad oauth2grant list --clientId $sp --output json | jq -r '.[] | select(.resourceId == "'"$graph"'") | {scope: .scope, resourceId: .resourceId')
  permissions=("Sites.Read.All" "Contacts.Read" "User.Read.All" "Mail.Read" "Calendars.Read" "Group.Read.All")
  if [ -z "$graphPermissions" ]; then
    echo "No permissions for MS Graph configured yet. Setting..."
    o365 aad oauth2grant add --clientId $sp --resourceId $graph --scope "$(echo $permissions)"
    success "DONE"
  else
    scope=$(echo $graphPermissions | jq -r '.scope')
    echo "Existing permissions found: $scope. Updating..."
    for permission in "${permissions[@]}"; do
      if [[ ! $scope == *"$permission"* ]]; then
        scope="$scope $permission"
      fi
    done
    
    graphPermissionsResourceId=$(echo $graphPermissions | jq -r '.resourceId')
    echo "Updating permissions to $scope..."
    o365 spo oauth2grant set --grantId $resourceId --scope "$scope"
    success "DONE"
  fi

  # TODO
  # disable quick launch for the portal
  # deploy.ps1:100

  theme=$(o365 spo propertybag get --webUrl $portalUrl --key ThemePrimary --output json)

  id=$(echo $app | jq -r '.ID')
  installedVersion=$(echo $app | jq -r '.InstalledVersion')
  if [ -z $installedVersion ]; then
    o365 spo app upgrade --id $id --siteUrl $portalUrl
  fi
fi

