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



echo "Disabling quick launch for the portal..."
o365 spo web set --webUrl $portalUrl --quickLaunchEnabled false
success "DONE"

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
  o365 spo storageentity set --appCatalogUrl $appCatalogUrl --key "PnP-Portal-AlphaVantage-API-Key" --value "$stockAPIKey" --description "API Key for Alpha Advantage REST Stock service"
  success "DONE"
else
  warning "stockAPIKey not specified. Skipping"
fi