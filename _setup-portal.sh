msg "Provisioning portal site at $portalUrl..."
site=$(o365 spo site get --url $portalUrl --output json || true)
if $(isError "$site"); then
  o365 spo site add --type CommunicationSite --url $portalUrl --title "$company" --description 'PnP SP Starket Kit Hub' >/dev/null
  success 'DONE'
else
  warning 'EXISTS'
fi

sub '- Applying theme...'
o365 spo theme apply --name "$company HR" --webUrl $portalUrl >/dev/null
success 'DONE'

sub '- Configuring hub site...'
hubsiteId=$(o365 spo hubsite list --output json | jq -r '.[] | select(.SiteUrl == "'"$portalUrl"'") | .ID')
if [ -z "$hubsiteId" ]; then
  out=$(o365 spo hubsite register --url $portalUrl --output json || true)
  if $(isError "$out"); then
    error 'ERROR'
    errorMessage "$out"
    exit 1
  fi
  hubsiteId=$(echo "$out" | jq -r '.ID')
  success 'DONE'
else
  warning 'EXISTS'
fi

sub '- Configuring logo...'
o365 spo web set --webUrl $portalUrl --siteLogoUrl $(echo $portalUrl)/SiteAssets/contoso_sitelogo.png
success 'DONE'

sub '- Provisioning site columns...\n'
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

for fieldInfo in "${fields[@]}"; do
  fieldId=$(echo "$fieldInfo" | grep -o '\sID="{[^}]*' | cut -d"{" -f2)
  sub "  - $fieldId..."
  field=$(o365 spo field get --webUrl $portalUrl --id $fieldId --output json || true)
  if $(isError "$field"); then
    o365 spo field add --webUrl $portalUrl --xml "$fieldInfo" >/dev/null
    success 'DONE'
  else
    warning 'EXISTS'
  fi
done

sub '- Provisioning content types...\n'
contentTypes=(
  'ID:"0x01007926A45D687BA842B947286090B8F67D" Name:"PnP Alert"',
  'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Name:"PnPTile"'
)

for contentTypeInfo in "${contentTypes[@]}"; do
  contentTypeId=$(getPropertyValue "$contentTypeInfo" "ID")
  sub "  - $contentTypeId..."
  contentType=$(o365 spo contenttype get --webUrl $portalUrl --id $contentTypeId --output json || true)
  if $(isError "$contentType"); then
    contentTypeName=$(getPropertyValue "$contentTypeInfo" "Name")
    o365 spo contenttype add --webUrl $portalUrl --id $contentTypeId --name "$contentTypeName" --group 'PnP Content Types'
    success 'DONE'
  else
    warning 'EXISTS'
  fi
done

sub '- Adding fields to content types...\n'
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

for ctField in "${contentTypesFields[@]}"; do
  contentTypeId=$(getPropertyValue "$ctField" "ID")
  fieldId=$(getPropertyValue "$ctField" "Field")
  required=$(if [[ $ctField = *"required:"* ]]; then echo "--required true"; else echo ""; fi)
  sub "  - Adding field $fieldId to content type $contentTypeId..."
  o365 spo contenttype field set --webUrl $portalUrl --contentTypeId $contentTypeId --fieldId $fieldId $required
  success 'DONE'
done

sub '- Creating pages...\n'
pages=(
  'Name:"home.aspx" Title:"Home" Layout:"Home" PromoteAsNewsArticle:"false"',
  'Name:"About-Us.aspx" Title:"About Us" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"HR.aspx" Title:"HR" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"People-Directory.aspx" Title:"People Directory" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"My-Profile.aspx" Title:"My Profile" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"Travel-Instructions.aspx" Title:"Travel Instructions" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"Financial-Results.aspx" Title:"Financial Results" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"FAQ.aspx" Title:"FAQ" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"Training.aspx" Title:"Training" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"Support.aspx" Title:"Support" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"Feedback.aspx" Title:"Feedback" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"Personal.aspx" Title:"Personal" Layout:"Article" PromoteAsNewsArticle:"false"',
  'Name:"Meeting-on-Marketing-In-Non-English-Speaking-Markets-This-Friday.aspx" Title:"Meeting on Marketing In Non-English-Speaking Markets This Friday" Layout:"Article" PromoteAsNewsArticle:"true"',
  'Name:"Marketing-Lunch.aspx" Title:"Marketing lunch" Layout:"Article" PromoteAsNewsArticle:"true"',
  'Name:"New-International-Marketing-Initiatives.aspx" Title:"New International Marketing Initiatives" Layout:"Article" PromoteAsNewsArticle:"true"',
  'Name:"New-Portal.aspx" Title:"New intranet portal" Layout:"Article" PromoteAsNewsArticle:"true"',
)

for pageInfo in "${pages[@]}"; do
  pageName=$(getPropertyValue "$pageInfo" "Name")
  pageTitle=$(getPropertyValue "$pageInfo" "Title")
  layout=$(getPropertyValue "$pageInfo" "Layout")
  promoteAsNews=$(getPropertyValue "$pageInfo" "PromoteAsNewsArticle")
  promote=$(if $promoteAsNews = 'true'; then echo "--promoteAs NewsPage"; else echo ""; fi)
  sub "  - $pageName..."
  page=$(o365 spo page get --webUrl $portalUrl --name $pageName --output json || true)
  if ! isError "$page"; then
    warning 'EXISTS'
    warningMsg "  - Removing $pageName..."
    o365 spo page remove --webUrl $portalUrl --name $pageName --confirm
    success 'DONE'
    sub "  - Creating $pageName..."
  fi
  # TODO: remove layout type once we can provision sections, otherwise we end up with an empty page without sections to which we can't add web parts
  o365 spo page add --webUrl $portalUrl --name $pageName --title "$pageTitle" --layoutType $layout $promote --publish
  success 'DONE'
done

sub '- Configuring navigation...\n'
# remove old navigation nodes
navigationNodes=($(o365 spo navigation node list --webUrl $portalUrl --location TopNavigationBar --output json | jq '.[] | .Id'))
exists=${navigationNodes:-}
if [ ! -z ${exists} ]; then
  for node in "${navigationNodes[@]}"; do
    warningMsg "  - Removing node $node..."
    o365 spo navigation node remove --webUrl $portalUrl --location TopNavigationBar --id $node --confirm
    success 'DONE'
  done
fi
# create new navigation nodes
navigationNodes=(
  'Title:"Personal" Url:"SitePages/Personal.aspx"',
  'Title:"Organization" Url:"SitePages/Home.aspx"',
  'Title:"Departments" Url:" "'
)
departmentsNodes=(
  "Title:\"Human Resources\" Url:\"$hrUrl\"",
  "Title:\"Marketing\" Url:\"$marketingUrl\""
)
for node in "${navigationNodes[@]}"; do
  nodeTitle=$(getPropertyValue "$node" "Title")
  nodeUrl=$(getPropertyValue "$node" "Url")
  sub "  - $nodeTitle..."
  result=$(o365 spo navigation node add --webUrl $portalUrl --location TopNavigationBar --title "$nodeTitle" --url "$nodeUrl" --output json)
  success 'DONE'
done
departmentsNodeId=$(echo $result | jq '.Id')

for node in "${departmentsNodes[@]}"; do
  nodeTitle=$(getPropertyValue "$node" "Title")
  nodeUrl=$(getPropertyValue "$node" "Url")
  sub "    - $nodeTitle..."
  o365 spo navigation node add --webUrl $portalUrl --parentNodeId $departmentsNodeId --title "$nodeTitle" --url "$nodeUrl" >/dev/null
  success 'DONE'
done

sub '- Provisioning lists...\n'
# events
sub '  - Events...'
list=$(o365 spo list get --webUrl $portalUrl --title Events --output json || true)
if $(isError "$list"); then
  o365 spo list add --webUrl $portalUrl --title Events --baseTemplate Events \
    --templateFeatureId 00bfea71-ec85-4903-972d-ebe475780106 \
    --contentTypesEnabled >/dev/null
  success 'DONE'
else
  warning 'EXISTS'
fi

sub '    - List items...\n'
addOrUpdateListItem $portalUrl Events 'SharePoint Conference North America' \
  --Title 'SharePoint Conference North America' \
  --fAllDayEvent true \
  --EventDate '2018-05-21 00:00:00' \
  --EndDate '2018-05-23 23:59:00'
addOrUpdateListItem $portalUrl Events 'European Collaboration Summit' \
  --Title 'European Collaboration Summit' \
  --fAllDayEvent true \
  --EventDate '2018-05-28 00:00:00' \
  --EndDate '2018-05-30 23:59:00'
addOrUpdateListItem $portalUrl Events 'Microsoft Ignite' \
  --Title 'Microsoft Ignite' \
  --fAllDayEvent true \
  --EventDate '2018-09-24 00:00:00' \
  --EndDate '2018-09-28 23:59:00'
addOrUpdateListItem $portalUrl Events 'European SharePoint Conference' \
  --Title 'European SharePoint Conference' \
  --fAllDayEvent true \
  --EventDate '2018-11-26 00:00:00' \
  --EndDate '2018-11-29 23:59:00'
addOrUpdateListItem $portalUrl Events 'SharePoint Conference' \
  --Title 'SharePoint Conference' \
  --fAllDayEvent true \
  --EventDate '2019-05-21 00:00:00' \
  --EndDate '2019-05-23 23:59:00'
addOrUpdateListItem $portalUrl Events 'European Collaboration Summit' \
  --Title 'SharePoint Conference' \
  --fAllDayEvent true \
  --EventDate '2019-05-27 00:00:00' \
  --EndDate '2019-05-29 23:59:00'
# /events
# alerts
sub '  - Alerts...'
list=$(o365 spo list get --webUrl $portalUrl --title Alerts --output json || true)
if $(isError "$list"); then
  o365 spo list add --webUrl $portalUrl --title Alerts --baseTemplate GenericList \
    --templateFeatureId 00bfea71-de22-43b2-a848-c05709900100 \
    --contentTypesEnabled >/dev/null
  success 'DONE'
else
  warning 'EXISTS'
fi
sub '    - Adding PnP Alert content type...'
contentType=$(o365 spo list contenttype list --webUrl $portalUrl \
  --listTitle Alerts --output json | \
  jq -r '.[] | select(.StringId | startswith("0x01007926A45D687BA842B947286090B8F67D")) | .StringId')
if [ -z "$contentType" ]; then
  o365 spo list contenttype add --webUrl $portalUrl --listTitle Alerts \
    --contentTypeId 0x01007926A45D687BA842B947286090B8F67D >/dev/null
  success 'DONE'
else
  warning 'EXISTS'
fi
# /alerts

success 'DONE'
echo