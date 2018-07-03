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

tenantUrl=
prefix=""
skipSolutionDeployment=false
skipSiteCreation=false
stockAPIKey=""
company="Contoso"
weatherCity="Seattle"
stockSymbol="MSFT"

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
    *)
      echo "Invalid argument $1"
      exit 1
  esac
  shift
done

if [ -z "$tenantUrl" ]; then
  echo "Please specify tenant URL"
  exit 1
fi

if [ ! $skipSiteCreation = true ]; then
  . ./_create-hierarchy.sh
fi
. ./_provision-solution-prerequisites.sh
if [ ! $skipSolutionDeployment = true ]; then
  . ./_deploy-solution-package.sh
fi
. ./_setup-portal.sh
. ./_setup-department-sites.sh