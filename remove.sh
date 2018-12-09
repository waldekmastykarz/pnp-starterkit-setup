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
. ./_functions.sh

# Prerequisites
msg 'Checking prerequisites...'

set +e
_=$(command -v o365);
if [ "$?" != "0" ]; then
  error 'ERROR'
  echo
  echo "You don't seem to have the Office 365 CLI installed."
  echo "Install it by executing 'npm i -g @pnp/office365-cli'"
  echo "More information: https://aka.ms/o365cli"
  exit 127
fi;

_=$(command -v jq);
if [ "$?" != "0" ]; then
  error 'ERROR'
  echo
  echo "You don't seem to have jq installed."
  echo "Install it from https://stedolan.github.io/jq/"
  exit 127
fi;
set -e
success 'DONE'

# default args values
tenantUrl=
prefix=""
company="Contoso"

# script arguments
while [ $# -gt 0 ]; do
  case $1 in
    -t|--tenantUrl)
      shift
      tenantUrl=$1
      ;;
    -p|--prefix)
      shift
      prefix=$1
      ;;
    -c|--company)
      shift
      company=$1
      ;;
    -h|--help)
      help
      exit
      ;;
    *)
      error "Invalid argument $1"
      exit 1
  esac
  shift
done

if [ -z "$tenantUrl" ]; then
  error 'Please specify tenant URL'
  echo
  help
  exit 1
fi

msg 'Retrieving tenant app catalog URL...'
portalUrl=$tenantUrl/sites/$(echo $prefix)portal
hrUrl=$tenantUrl/sites/$(echo $prefix)hr
marketingUrl=$tenantUrl/sites/$(echo $prefix)marketing
appCatalogUrl=$(o365 spo tenant appcatalogurl get)
if [ -z "$appCatalogUrl" ]; then
  error "Couldn't retrieve tenant app catalog"
  exit 1
fi
success 'DONE'

echo
msg 'Removing the SP Starter Kit...\n'
echo

manualRequired=false

# Sites
sites=($hrUrl $marketingUrl $portalUrl)
for siteUrl in "${sites[@]}"; do
  sub "- Remove site $siteUrl..."
  site=$(o365 spo site get --url $siteUrl --output json || true)
  if $(isError "$site"); then
    success 'NOT FOUND'
  else
    manualRequired=true
    warning 'MANUALLY'
  fi
done
# / Sites

# Term group
sub '- Remove PnPTermSets Term Group...'
termGroups=$(o365 spo term group list --output json | jq '.[] | select(.Name == "PnPTermSets") | .Name')
if [ -z "$termGroups" ]; then
  success 'NOT FOUND'
else
  manualRequired=true
  warning 'MANUALLY'
fi
# / Term group

sub '- Removing the sharepoint-starter-kit.sppkg solution package from the tenant app catalog...'
app=$(o365 spo app get --name "sharepoint-starter-kit.sppkg" --output json || true)
if ! isError "$app"; then
  appId=$(echo $app | jq -r '.ID')
  o365 spo app remove --id $appId --confirm
  success 'DONE'
else
  success 'NOT FOUND'
fi

# API permissions
sub '- Removing API permissions...\n'
permissions=("Sites.Read.All" "Contacts.Read" "User.Read.All" "Mail.Read" "Calendars.ReadWrite" "Group.ReadWrite.All")
grants=$(o365 spo sp grant list --output json)
for permission in "${permissions[@]}"; do
  sub "  - $permission..."
  grantId=$(echo $grants | jq -r '.[] | select(.Resource == "Microsoft Graph" and .Scope == "'"$permission"'") | .ObjectId')
  if [ -z "$grantId" ]; then
    success 'NOT FOUND'
  else
    o365 spo sp grant revoke --grantId $grantId
    success 'DONE'
  fi
done
# / API permissions

# Themes
sub '- Removing themes...\n'
themes=$(o365 spo theme list --output json)
themeNames=("HR" "Marketing" "Sales" "Technologies")
for themeName in "${themeNames[@]}"; do
  fullThemeName="$company $themeName"
  sub "  - $fullThemeName..."
  theme=$(echo $themes | jq -r '.[] | select(.name == "'"$fullThemeName"'") | .name')
  if [ -z "$theme" ]; then
    success 'NOT FOUND'
  else
    o365 spo theme remove --name "$fullThemeName" --confirm
    success 'DONE'
  fi
done
# / Themes

# Site scripts
sub '- Removing site scripts...\n'
scripts=$(o365 spo sitescript list --output json)
siteScripts=('Team Site' 'Communication Site')
for siteScript in "${siteScripts[@]}"; do
  scriptName="$company $siteScript"
  sub "  - $scriptName..."
  scriptId=$(echo $scripts | jq -r '.[] | select(.Title == "'"$scriptName"'") | .Id')
  if [ -z "$scriptId" ]; then
    success 'NOT FOUND'
  else
    o365 spo sitescript remove --id $scriptId --confirm
    success 'DONE'
  fi
done
# / Site scripts

# Site designs
sub '- Removing site designs...\n'
designs=$(o365 spo sitedesign list --output json)
siteDesigns=('Team Site' 'Communication Site')
for siteDesign in "${siteDesigns[@]}"; do
  designName="$company $siteDesign"
  sub "  - $designName..."
  designId=$(echo $designs | jq -r '.[] | select(.Title == "'"$designName"'") | .Id')
  if [ -z "$designId" ]; then
    success 'NOT FOUND'
  else
    o365 spo sitedesign remove --id $designId --confirm
    success 'DONE'
  fi
done
# / Site designs

echo
success "SP Starter Kit has been successfully removed"
echo

if [ "$manualRequired" = true ]; then
  warning "Some resources couldn't have been removed automatically"
  warning "and need to be removed manually."
  warning "See the log above for more details."
fi