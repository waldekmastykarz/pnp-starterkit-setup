portalUrl=$tenantUrl/sites/$(echo $prefix)portal

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