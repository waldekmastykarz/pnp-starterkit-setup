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

fields=(
  '`<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>`'
  '`<Field Type="URL" DisplayName="More information link" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Hyperlink" Group="PnP Columns" ID="{6085e32a-339b-4da7-ab6d-c1e013e5ab27}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertMoreInformation" Name="PnPAlertMoreInformation"></Field>`'
  '`<Field Type="Note" DisplayName="Alert message" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" NumLines="6" RichText="FALSE" Sortable="FALSE" Group="PnP Columns" ID="{f056406b-b46b-4a94-8503-361de4ca2752}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertMessage" Name="PnPAlertMessage"></Field>`'
  '`<Field Type="Choice" DisplayName="Alert type" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="RadioButtons" FillInChoice="FALSE" Group="PnP Columns" ID="{ebe7e498-44ff-43da-a7e5-99b444f656a5}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertType" Name="PnPAlertType"><Default>Information</Default><CHOICES><CHOICE>Information</CHOICE><CHOICE>Urgent</CHOICE></CHOICES></Field>`'
  '`<Field Type="DateTime" DisplayName="End date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{b0d8c6ed-2487-43e7-a716-bf274f0d5e09}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertEndDateTime" Name="PnPAlertEndDateTime"><Default>[today]</Default></Field>`'
  '`<Field Type="Choice" DisplayName="PnP Url Target" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" Group="PnP Columns" ID="{4ad64f28-1772-492d-bde4-998a08f8a7ae}" SourceID="{2765be97-bf5e-434d-b563-9f1b907b5397}" StaticName="PnPUrlTarget" Name="PnPUrlTarget"><CHOICES><CHOICE>Current window</CHOICE><CHOICE>New window</CHOICE></CHOICES></Field>`'
  '`<Field Type="Choice" DisplayName="PnP Icon Name" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="TRUE" Group="PnP Columns" ID="{a374d13d-040b-4104-9973-edfdca1e3fc1}" SourceID="{2765be97-bf5e-434d-b563-9f1b907b5397}" StaticName="PnPIconName" Name="PnPIconName"><CHOICES><CHOICE>12PointStar</CHOICE><CHOICE>6PointStar</CHOICE><CHOICE>AADLogo</CHOICE><CHOICE>Accept</CHOICE><CHOICE>Accounts</CHOICE></CHOICES></Field>`'
  '`<Field Type="Text" DisplayName="PnP Url" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group="PnP Fields" ID="{9913d58a-9a75-41fc-86aa-f6de1d9328ba}" SourceID="{2765be97-bf5e-434d-b563-9f1b907b5397}" StaticName="PnPUrl" Name="PnPUrl"></Field>`'
  '`<Field Type="Note" DisplayName="PnP Description" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" NumLines="6" RichText="FALSE" Sortable="FALSE" Group="PnP Columns" ID="{a67b73bf-acb0-4517-93e6-91585316ec49}" SourceID="{2765be97-bf5e-434d-b563-9f1b907b5397}" StaticName="PnPDescription" Name="PnPDescription"></Field>`'
)

echo "Provisioning site columns..."
for fieldInfo in "${fields[@]}"; do
  fieldId=$(echo "$fieldInfo" | grep -o '\sID="{[^}]*' | cut -d"{" -f2)
  echo "  Provisioning $fieldId..."
  field=$(o365 spo field get --webUrl $portalUrl --id $fieldId --output json || true)
  if $(isError "$field"); then
    echo "    Creating site column..."
    o365 spo field add --webUrl $portalUrl --xml "$fieldInfo"
    success "    DONE"
  else
    warning "    Site column already exists"
  fi
done
success "DONE"

contentTypes=(
  'ID:"0x01007926A45D687BA842B947286090B8F67D" Name:"PnP Alert"',
  'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Name:"PnPTile"'
)

echo "Provisioning content types..."
for contentTypeInfo in "${contentTypes[@]}"; do
  contentTypeId=$(getPropertyValue "$contentTypeInfo" "ID")
  echo "  Provisioning content type $contentTypeId..."
  contentType=$(o365 spo contenttype get --webUrl $portalUrl --id $contentTypeId --output json || true)
  if $(isError "$contentType"); then
    echo "    Creating content type..."
    contentTypeName=$(getPropertyValue "$contentTypeInfo" "Name")
    o365 spo contenttype add --webUrl $portalUrl --id $contentTypeId --name "$contentTypeName" --group 'PnP Content Types'
    success "    DONE"
  else
    warning "    Content type already exists"
  fi
done
success "DONE"

# The Name argument is purely informative so that you can easily see which field it is
contentTypesFields=(
  'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"ebe7e498-44ff-43da-a7e5-99b444f656a5" Name:"PnPAlertType" required:"true"',
  'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"f056406b-b46b-4a94-8503-361de4ca2752" Name:"PnPAlertMessage" required:"true"',
  'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"5ee2dd25-d941-455a-9bdb-7f2c54aed11b" Name:"PnPAlertStartDateTime" required:"true"',
  'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"b0d8c6ed-2487-43e7-a716-bf274f0d5e09" Name:"PnPAlertEndDateTime" required:"true"',
  'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"6085e32a-339b-4da7-ab6d-c1e013e5ab27" Name:"PnPAlertMoreInformation"',

  'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Field:"a67b73bf-acb0-4517-93e6-91585316ec49" Name:"PnPDescription"',
  'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Field:"a374d13d-040b-4104-9973-edfdca1e3fc1" Name:"PnPIconName"',
  'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Field:"9913d58a-9a75-41fc-86aa-f6de1d9328ba" Name:"PnPUrl"',
  'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Field:"4ad64f28-1772-492d-bde4-998a08f8a7ae" Name:"PnPUrlTarget"'
)

echo "Adding fields to content types..."
for ctField in "${contentTypesFields[@]}"; do
  contentTypeId=$(getPropertyValue "$ctField" "ID")
  fieldId=$(getPropertyValue "$ctField" "Field")
  required=$(if [[ $ctField = *"required:"* ]]; then echo "--required true"; else echo ""; fi)
  echo "  Adding field $fieldId to content type $contentTypeId..."
  o365 spo contenttype field set --webUrl $portalUrl --contentTypeId $contentTypeId --fieldId $fieldId $required
  success "  DONE"
done
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
o365 spo page add --webUrl $portalUrl --name $pageName
success "  DONE"
echo "  Adding sections..."
echo "    #1"
# TODO: o365 spo page section add --webUrl $portalUrl --name $pageName --sectionTemplate TwoColumnLeft --order 1

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