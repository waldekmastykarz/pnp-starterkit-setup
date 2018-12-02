echo "Retrieving portal..."
siteJson=$(o365 spo site get --url $portalUrl --output json)
if [ -z "$siteJson" ]; then
  error "Portal site at URL $portalUrl not found"
  exit 1
fi

siteId=$(echo $siteJson | jq -r '.Id')
success "DONE"

echo "Setting PnPTilesList tenant property..."
o365 spo storageentity set --appCatalogUrl $appCatalogUrl --key PnPTilesList-$(echo $siteJson) --value $portalUrl/Lists/PnPTiles
success "DONE"





# Provision pages

## Home page

pageName="home.aspx"
echo "Provisioning page $pageName..."
page=$(o365 spo page get --webUrl $portalUrl --name $pageName --output json || true)
if ! isError "$page"; then
  warning "  Page $pageName already exists. Removing..."
  o365 spo page remove --webUrl $portalUrl --name $pageName --confirm
  success "  DONE"
fi
echo "  Creating page..."
# TODO: remove layout type once we can provision sections, otherwise we end up with an empty page without sections to which we can't add web parts
o365 spo page add --webUrl $portalUrl --name $pageName --layoutType Home
success "  DONE"
echo "  Adding sections..."
echo "    #1"
# TODO: o365 spo page section add --webUrl $portalUrl --name $pageName --sectionTemplate TwoColumnLeft --order 1
echo "      Adding web parts..."
echo "        Hero"
o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName --standardWebPart Hero --webPartPropertiesFile portal-home-01-1-1-hero.json
success "        DONE"
success "      DONE"
success "    DONE"


# Navigation requires pages to be provisioned first
exit 1
echo "Configuring navigation..."
navigationNodes=($(o365 spo navigation node list --webUrl $portalUrl --location TopNavigationBar --output json | jq '.[] | .Id'))
for node in "${navigationNodes}"; do
  echo "  Removing node $node..."
  o365 spo navigation node remove --webUrl $portalUrl --location TopNavigationBar --id $node --confirm
  success "  DONE"
done
echo "  Creating node Personal..."
o365 spo navigation node add --webUrl $portalUrl --location TopNavigationBar --title Personal --url SitePages/Personal.aspx
success "  DONE"
echo "  Creating node Organization..."
o365 spo navigation node add --webUrl $portalUrl --location TopNavigationBar --title Organization --url SitePages/Home.aspx
success "  DONE"
echo "  Creating node Departments..."
o365 spo navigation node add --webUrl $portalUrl --location TopNavigationBar --title Departments --url ' '
success "  DONE"
success "DONE"

# TODO: provision lists
# TODO: provision custom actions
# TODO: provision files
# TODO: install .sppkg