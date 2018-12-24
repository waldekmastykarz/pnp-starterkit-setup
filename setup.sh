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
skipSolutionDeployment=false
skipSiteCreation=false
stockAPIKey=""
company="Contoso"
weatherCity="Seattle, WA"
stockSymbol="MSFT"
checkPoint=0

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
    --skipSolutionDeployment)
      skipSolutionDeployment=true
      ;;
    --skipSiteCreation)
      skipSiteCreation=true
      ;;
    --stockAPIKey)
      shift
      stockAPIKey=$1
      ;;
    -c|--company)
      shift
      company=$1
      ;;
    -w|--weatherCity)
      shift
      weatherCity=$1
      ;;
    --stockSymbol)
      shift
      stockSymbol=$1
      ;;
    --checkPoint)
      shift
      checkPoint=$1
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

# show check point information to allow the user to resume the script
# from the last successful state
trap "checkPoint" ERR

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
msg 'Provisioning the SP Starter Kit...\n'
echo

if (( $checkPoint < 100 )); then
  . ./_setup-tenant.sh
fi
if (( $checkPoint < 200 )); then
  . ./_setup-taxonomy.sh
fi
. ./_setup-portal.sh
if (( $checkPoint < 400 )); then
  . ./_setup-hr.sh
fi
if (( $checkPoint < 500 )); then
  . ./_setup-marketing.sh
fi

if (( $checkPoint < 600 )); then
  sub "- Setting stockAPIKey $stockAPIKey..."
  if [ ! -z "$stockAPIKey" ]; then
    o365 spo storageentity set --appCatalogUrl $appCatalogUrl \
      --key "PnP-Portal-AlphaVantage-API-Key" --value "$stockAPIKey" \
      --description "API Key for Alpha Advantage REST Stock service"
    success 'DONE'
  else
    warning 'SKIPPED'
  fi
  checkPoint=600
fi

echo
success "SP Starter Kit has been successfully provisioned to $portalUrl"