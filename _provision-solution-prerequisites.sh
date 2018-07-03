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
  warning "Origin already exists"
fi

echo "Configuring web API permissions..."
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
graphPermissions=$(o365 aad oauth2grant list --clientId $sp --output json | jq -r '.[] | select(.resourceId == "'"$graph"'") | {scope: .scope, objectId: .objectId'})
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

  graphPermissionsObjectId=$(echo $graphPermissions | jq -r '.objectId')
  echo "Updating permissions to $scope..."
  o365 aad oauth2grant set --grantId $graphPermissionsObjectId --scope "$scope"
  success "DONE"
fi

# TODO
# disable quick launch for the portal
# deploy.ps1:100

currentTheme=$(o365 spo propertybag get --webUrl $portalUrl --key ThemePrimary)
theme=$(cat theme.json | jq '.themePrimary')
if [ ! "$currentTheme" = "$theme" ]; then
  echo "Setting theme..."
  o365 spo theme set --name "Contoso Portal" --filePath ./theme.json
  success "DONE"
  echo "Applying theme..."
  o365 spo theme apply --name "Contoso Portal" --webUrl $portalUrl
  success "DONE"
else
  warning "Theme already set"
fi

if [ ! -z "$stockAPIKey" ]; then
  echo "Setting stockAPIKey $stockAPIKey..."
  # TODO don't require to specify the app catalog URL
  o365 spo storageentity set --key "PnP-Portal-AlphaVantage-API-Key" --value "$stockAPIKey" --description "API Key for Alpha Advantage REST Stock service"
  success "DONE"
else
  warning "stockAPIKey not specified. Skipping"
fi