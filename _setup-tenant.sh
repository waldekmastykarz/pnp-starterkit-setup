### API permissions
msg 'Configuring API permissions...\n'
permissions=("Sites.Read.All" "Contacts.Read" "User.Read.All" "Mail.Read" "Calendars.ReadWrite" "Group.ReadWrite.All")
for permission in "${permissions[@]}"; do
  sub "- Adding permission $permission to Microsoft Graph..."
  res=$(o365 spo sp grant add --resource 'Microsoft Graph' --scope $permission --output json || true)
  if $(isError "$res"); then
    if [[ $res == *"already exists"* ]]; then
      warning 'EXISTS'
    else
      errorMessage "$res"
      exit 1
    fi
  else
    success 'DONE'
  fi
done
success 'DONE'
echo
### /API permissions

### solution package
if [ ! $skipSolutionDeployment = true ]; then
  msg 'Deploying solution...\n'

  app=$(o365 spo app get --name "sharepoint-starter-kit.sppkg" --output json || true)
  if ! isError "$app"; then
    warningMsg '- Solution package already exists. Removing...'
    appId=$(echo $app | jq -r '.ID')
    o365 spo app remove --id $appId --confirm
    success 'DONE'
  fi

  sub '- Adding solution package to tenant app catalog...'
  o365 spo app add --filePath ./sharepoint-starter-kit.sppkg >/dev/null
  success 'DONE'

  sub '- Deploying solution package...'
  o365 spo app deploy --name sharepoint-starter-kit.sppkg --skipFeatureDeployment
  success 'DONE'
  success 'DONE'
  echo
fi
### /solution package

### themes
msg 'Registering themes...\n'
themes=(
  'Name:"hr" Title:"HR"',
  'Name:"marketing" Title:"Marketing"',
  'Name:"sales" Title:"Sales"',
  'Name:"technologies" Title:"Technologies"'
)
for themeInfo in "${themes[@]}"; do
  themeName=$(getPropertyValue "$themeInfo" "Name")
  themeTitle=$(getPropertyValue "$themeInfo" "Title")
  themeFilePath="./resources/themes/$themeName.json"
  themeFullTitle="$company $themeTitle"
  sub "- Registering theme $themeFullTitle..."
  o365 spo theme set --name "$themeFullTitle" --filePath "$themeFilePath" >/dev/null
  success 'DONE'
done
success 'DONE'
echo
### /themes

### site scripts
msg 'Setting site scripts...\n'
siteScriptsJson=$(o365 spo sitescript list --output json)
teamSiteScriptName="$company Team Site"
teamSiteScriptId=$(echo $siteScriptsJson | jq -r '.[] | select(.Title == "'"$teamSiteScriptName"'") | .Id')
if [ ! -z "$teamSiteScriptId" ]; then
  warningMsg "- Site script '$teamSiteScriptName' already exists. Removing..."
  o365 spo sitescript remove --id $teamSiteScriptId --confirm
  success 'DONE'
fi
sub '- Setting team site site script...'
teamSiteScript=$(cat ./resources/sitescripts/collabteamsite.json | jq -c '.')
teamSiteScriptId=$(o365 spo sitescript add --title "$teamSiteScriptName" --content "'"$teamSiteScript"'" --output json | jq -r '.Id')
success 'DONE'

commSiteScriptName="$company Communication Site"
commSiteScriptId=$(echo $siteScriptsJson | jq -r '.[] | select(.Title == "'"$commSiteScriptName"'") | .Id')
if [ ! -z "$commSiteScriptId" ]; then
  warningMsg "- Site script '$commSiteScriptName' already exists. Removing..."
  o365 spo sitescript remove --id $commSiteScriptId --confirm
  success 'DONE'
fi
sub '- Setting communication site site script...'
commSiteScript=$(cat ./resources/sitescripts/collabcommunicationsite.json | jq -c '.')
commSiteScriptId=$(o365 spo sitescript add --title "$commSiteScriptName" --content "'"$commSiteScript"'" --output json | jq -r '.Id')
success 'DONE'

success 'DONE'
echo
### /site scripts

### site designs
msg 'Setting site designs...\n'

siteDesignsJson=$(o365 spo sitedesign list --output json)
teamSiteDesignName="$company Team Site"
teamSiteDesignId=$(echo $siteDesignsJson | jq -r '.[] | select(.Title == "'"$teamSiteDesignName"'") | .Id')
if [ ! -z "$teamSiteDesignId" ]; then
  warningMsg "- Site design '$teamSiteDesignName' already exists. Removing..."
  o365 spo sitedesign remove --id $teamSiteDesignId --confirm
  success 'DONE'
fi
sub "- Setting '$teamSiteDesignName' site design..."
o365 spo sitedesign add --title "$teamSiteDesignName" --webTemplate TeamSite --siteScripts "$teamSiteScriptId" >/dev/null
success 'DONE'

commSiteDesignName="$company Communication Site"
commSiteDesignId=$(echo $siteDesignsJson | jq -r '.[] | select(.Title == "'"$commSiteDesignName"'") | .Id')
if [ ! -z "$commSiteDesignId" ]; then
  warningMsg "- Site design '$commSiteDesignName' already exists. Removing..."
  o365 spo sitedesign remove --id $commSiteDesignId --confirm
  success 'DONE'
fi
sub "- Setting '$commSiteDesignName' site design..."
result=$(o365 spo sitedesign add --title "$commSiteDesignName" --webTemplate CommunicationSite --siteScripts "$commSiteScriptId" --output json)
success 'DONE'

success 'DONE'
echo
### /site designs