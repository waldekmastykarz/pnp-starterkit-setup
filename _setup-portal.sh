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

siteScriptsJson=$(o365 spo sitescript list --output json)

echo "Setting team site site script..."
teamSiteScriptName="$company Team Site"
teamSiteScriptId=$(echo $siteScriptsJson | jq -r '.[] | select(.Title == "'"$teamSiteScriptName"'") | .Id')
if [ ! -z "$teamSiteScriptId" ]; then
    warning "Site script '$teamSiteScriptName' already exists. Removing..."
    o365 spo sitescript remove --id $teamSiteScriptId --confirm
    success "DONE"
fi
teamSiteScript=$(cat collabteamsite.json | jq -c '.')
teamSiteScriptId=$(o365 spo sitescript add --title $teamSiteScriptName --content "'"$teamSiteScript"'" --output json | jq -r '.Id')
success "DONE"

echo "Setting communication site site script..."
commSiteScriptName="$company Communication Site"
commSiteScriptId=$(echo $siteScriptsJson | jq -r '.[] | select(.Title == "'"$commSiteScriptName"'") | .Id')
if [ ! -z "$commSiteScriptId" ]; then
    warning "Site script '$commSiteScriptName' already exists. Removing..."
    o365 spo sitescript remove --id $commSiteScriptId --confirm
    success "DONE"
fi
commSiteScript=$(cat collabcommunicationsite.json | jq -c '.')
commSiteScriptId=$(o365 spo sitescript add --title "$commSiteScriptName" --content "'"$commSiteScript"'" --output json | jq -r '.Id')
success "DONE"

siteDesignsJson=$(o365 spo sitedesign list --output json)

echo "Setting team site site design..."
teamSiteDesignName="$company Team Site"
teamSiteDesignId=$(echo $siteDesignsJson | jq -r '.[] | select(.Title == "'"$teamSiteDesignName"'") | .Id')
if [ ! -z "$teamSiteDesignId" ]; then
    warning "Site design '$teamSiteDesignName' already exists. Removing..."
    o365 spo sitedesign remove --id $teamSiteDesignId --confirm
    success "DONE"
fi
o365 spo sitedesign add --title "$teamSiteDesignName" --webTemplate TeamSite --siteScripts "$teamSiteScriptId"
success "DONE"

echo "Setting communication site site design..."
commSiteDesignName="$company Communication Site"
commSiteDesignId=$(echo $siteDesignsJson | jq -r '.[] | select(.Title == "'"$commSiteDesignName"'") | .Id')
if [ ! -z "$commSiteDesignId" ]; then
    warning "Site design '$commSiteDesignName' already exists. Removing..."
    o365 spo sitedesign remove --id $commSiteDesignId --confirm
    success "DONE"
fi
o365 spo sitedesign add --title "$commSiteDesignName" --webTemplate CommunicationSite --siteScripts "$commSiteScriptId"
success "DONE"

# TODO: Set navigation
    #   <pnp:Navigation>
    #     <pnp:GlobalNavigation NavigationType="Structural">
    #       <pnp:StructuralNavigation RemoveExistingNodes="true">
    #         <pnp:NavigationNode Title="Personal" Url="SitePages/Personal.aspx"/>
    #         <pnp:NavigationNode Title="Organization" Url="SitePages/Home.aspx"/>
    #         <pnp:NavigationNode Title="Departments" Url="" />
    #       </pnp:StructuralNavigation>
    #     </pnp:GlobalNavigation>
    #   </pnp:Navigation>

# TODO: provision fields
# TODO: provision content types
# TODO: add fields to content types
# TODO: provision pages
# TODO: provision lists
# TODO: provision custom actions
# TODO: provision files
# TODO: install .sppkg