if (( $checkPoint < 500 )); then
  msg "Provisioning portal site at $portalUrl..."
fi

if (( $checkPoint >= 210 )); then
    msg "\n"
fi
if (( $checkPoint < 210 )); then
  site=$(o365 spo site get --url $portalUrl --output json || true)
  if $(isError "$site"); then
    o365 spo site add --type CommunicationSite --url $portalUrl --title "$company" --description 'PnP SP Starket Kit Hub' --lcid 1033 >/dev/null
    success 'DONE'
  else
    warning 'EXISTS'
  fi

  checkPoint=210
fi

if (( $checkPoint < 280 )); then
  # IDs required for provisioning web parts
  sub '- Retrieving site ID...'
  siteId=$(o365 spo site get --url $portalUrl --output json | jq -r '.Id')
  success 'DONE'

  sub '- Retrieving root web ID...'
  webId=$(o365 spo web get --webUrl $portalUrl --output json | jq -r '.Id')
  success 'DONE'

  sub '- Retrieving Site Pages list ID...'
  sitePagesListId=$(o365 spo list get --webUrl $portalUrl --title 'Site Pages' --output json | jq -r '.Id')
  success 'DONE'

  sub '- Retrieving Documents list ID...'
  documentsListId=$(o365 spo list get --webUrl $portalUrl --title 'Documents' --output json | jq -r '.Id')
  success 'DONE'
fi

# hubsiteId is also required for provisioning marketing and HR sites
if (( $checkPoint < 500 )); then
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

  if (( $checkPoint >= 300 )); then
    echo
  fi
fi

if (( $checkPoint < 215 )); then
  sub '- Applying theme...'
  o365 spo theme apply --name "$company HR" --webUrl $portalUrl >/dev/null
  success 'DONE'

  checkPoint=215
fi

if (( $checkPoint < 220 )); then
  sub '- Configuring logo...'
  o365 spo web set --webUrl $portalUrl --siteLogoUrl $(echo $portalUrl)/SiteAssets/contoso_sitelogo.png
  success 'DONE'

  checkPoint=220
fi

if (( $checkPoint < 230 )); then
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
    '`<Field Type="Choice" DisplayName="Link Group" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="TRUE" Group="PnP Columns" ID="{f5b23751-56d4-4ec3-adf2-7b080d834f74}" SourceID="{88a250c5-31b2-4d64-b047-482e2a6e5e7a}" StaticName="PnPPortalLinkGroup" Name="PnPPortalLinkGroup" CustomFormatter=""><CHOICES><CHOICE>Main Links</CHOICE></CHOICES></Field>`'
    '`<Field Type="URL" DisplayName="Link URL" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Hyperlink" Group="PnP Columns" ID="{c10389a0-8b29-4866-951f-3ad8e138db03}" SourceID="{88a250c5-31b2-4d64-b047-482e2a6e5e7a}" StaticName="PnPPortalLinkUrl" Name="PnPPortalLinkUrl" CustomFormatter=""></Field>`'
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

  checkPoint=230
fi

if (( $checkPoint < 240 )); then
  sub '- Provisioning content types...\n'
  contentTypes=(
    'ID:"0x01007926A45D687BA842B947286090B8F67D" Name:"PnP Alert"'
    'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Name:"PnPTile"'
    'ID:"0x0100580DB2292968A34EA3748511017A6DD2" Name:"PnPPortalLink"'
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

  checkPoint=240
fi

if (( $checkPoint < 250 )); then
  sub '- Adding fields to content types...\n'
  # The Name argument is purely informative so that you can easily see which field it is
  contentTypesFields=(
    'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"ebe7e498-44ff-43da-a7e5-99b444f656a5" Name:"PnPAlertType" required:"true"'
    'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"f056406b-b46b-4a94-8503-361de4ca2752" Name:"PnPAlertMessage" required:"true"'
    'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"5ee2dd25-d941-455a-9bdb-7f2c54aed11b" Name:"PnPAlertStartDateTime" required:"true"'
    'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"b0d8c6ed-2487-43e7-a716-bf274f0d5e09" Name:"PnPAlertEndDateTime" required:"true"'
    'ID:"0x01007926A45D687BA842B947286090B8F67D" Field:"6085e32a-339b-4da7-ab6d-c1e013e5ab27" Name:"PnPAlertMoreInformation"'

    'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Field:"a67b73bf-acb0-4517-93e6-91585316ec49" Name:"PnPDescription"'
    'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Field:"a374d13d-040b-4104-9973-edfdca1e3fc1" Name:"PnPIconName"'
    'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Field:"9913d58a-9a75-41fc-86aa-f6de1d9328ba" Name:"PnPUrl"'
    'ID:"0x0100FF0B2E33A3718B46A3909298D240FD92" Field:"4ad64f28-1772-492d-bde4-998a08f8a7ae" Name:"PnPUrlTarget"'

    'ID:"0x0100580DB2292968A34EA3748511017A6DD2" Field:"c10389a0-8b29-4866-951f-3ad8e138db03" Name:"PnPPortalLinkUrl"'
    'ID:"0x0100580DB2292968A34EA3748511017A6DD2" Field:"f5b23751-56d4-4ec3-adf2-7b080d834f74" Name:"PnPPortalLinkGroup"'
  )

  for ctField in "${contentTypesFields[@]}"; do
    contentTypeId=$(getPropertyValue "$ctField" "ID")
    fieldId=$(getPropertyValue "$ctField" "Field")
    required=$(if [[ $ctField = *"required:"* ]]; then echo "--required true"; else echo ""; fi)
    sub "  - Adding field $fieldId to content type $contentTypeId..."
    o365 spo contenttype field set --webUrl $portalUrl --contentTypeId $contentTypeId --fieldId $fieldId $required
    success 'DONE'
  done

  checkPoint=250
fi

if (( $checkPoint < 280 )); then
  sub '- Provisioning lists...\n'
    # events
    sub '  - Events...'
    list=$(o365 spo list get --webUrl $portalUrl --title Events --output json || true)
    if $(isError "$list"); then
      list=$(o365 spo list add --webUrl $portalUrl --title Events --baseTemplate Events \
        --templateFeatureId 00bfea71-ec85-4903-972d-ebe475780106 \
        --contentTypesEnabled --output json || true)
      if $(isError "$list"); then
        error 'ERROR'
        errorMessage "$list"
        exit 1
      fi
      eventsListId=$(echo $list | jq -r '.Id')
      success 'DONE'
    else
      eventsListId=$(echo $list | jq -r '.Id')
      warning 'EXISTS'
    fi

  if (( $checkPoint < 260 )); then
    sub '    - List items...\n'
    addOrUpdateListItem $portalUrl Events 'SharePoint Conference North America' \
      --fAllDayEvent true \
      --EventDate '2018-05-21 00:00:00' \
      --EndDate '2018-05-23 23:59:00'
    addOrUpdateListItem $portalUrl Events 'European Collaboration Summit' \
      --fAllDayEvent true \
      --EventDate '2018-05-28 00:00:00' \
      --EndDate '2018-05-30 23:59:00'
    addOrUpdateListItem $portalUrl Events 'Microsoft Ignite' \
      --fAllDayEvent true \
      --EventDate '2018-09-24 00:00:00' \
      --EndDate '2018-09-28 23:59:00'
    addOrUpdateListItem $portalUrl Events 'European SharePoint Conference' \
      --fAllDayEvent true \
      --EventDate '2018-11-26 00:00:00' \
      --EndDate '2018-11-29 23:59:00'
    addOrUpdateListItem $portalUrl Events 'SharePoint Conference' \
      --fAllDayEvent true \
      --EventDate '2019-05-21 00:00:00' \
      --EndDate '2019-05-23 23:59:00'
    addOrUpdateListItem $portalUrl Events 'European Collaboration Summit' \
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
    sub '    - Configuring All items view...'
    o365 spo list view set --webUrl $portalUrl --listTitle Alerts \
      --viewTitle 'All Items' \
      --ListViewXml '`<Query><OrderBy><FieldRef Name="ID" /></OrderBy></Query><ViewFields><FieldRef Name="LinkTitle" /><FieldRef Name="PnPAlertType" /><FieldRef Name="PnPAlertMessage" /><FieldRef Name="PnPAlertStartDateTime" /><FieldRef Name="PnPAlertEndDateTime" /><FieldRef Name="PnPAlertMoreInformation" /></ViewFields><RowLimit Paged="TRUE">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default="TRUE">main.xsl</XslLink><Toolbar Type="Standard"/>`'
    success 'DONE'
    # /alerts
  fi
  if (( $checkPoint < 280 )); then
    # Site Assets
    sub '  - Site Assets...'
    list=$(o365 spo list get --webUrl $portalUrl --title 'Site Assets' --output json || true)
    if $(isError "$list"); then
      list=$(o365 spo list add --webUrl $portalUrl --title 'SiteAssets' \
        --description 'Use this library to store files which are included on pages within this site, such as images on Wiki pages.' \
        --baseTemplate DocumentLibrary \
        --templateFeatureId 00bfea71-e717-4e80-aa17-d0c71b360101 \
        --contentTypesEnabled --output json || true)
      if $(isError "$list"); then
        error 'ERROR'
        errorMessage "$list"
        exit 1
      fi
      siteAssetsListId=$(echo $list | jq -r '.Id')
      o365 spo list set --webUrl $portalUrl --id $siteAssetsListId --title 'Site Assets'
      success 'DONE'
    else
      siteAssetsListId=$(echo $list | jq -r '.Id')
      warning 'EXISTS'
    fi
    # /Site Assets
  fi
  if (( $checkPoint < 260 )); then
    # PnP-PortalFooter-Links
    sub '  - PnP-PortalFooter-Links...'
    list=$(o365 spo list get --webUrl $portalUrl --title PnP-PortalFooter-Links --output json || true)
    if $(isError "$list"); then
      list=$(o365 spo list add --webUrl $portalUrl --title PnP-PortalFooter-Links --baseTemplate GenericList \
        --templateFeatureId 00bfea71-de22-43b2-a848-c05709900100 \
        --contentTypesEnabled --output json || true)
      if $(isError "$list"); then
        error 'ERROR'
        errorMessage "$list"
        exit 1
      fi
      listId=$(echo $list | jq -r '.Id')
      o365 spo list set --webUrl $portalUrl --id $listId --title PnP-PortalFooter-Links
      success 'DONE'
    else
      warning 'EXISTS'
    fi
    sub '    - Adding PnPPortalLink content type...'
    contentType=$(o365 spo list contenttype list --webUrl $portalUrl \
      --listTitle PnP-PortalFooter-Links --output json | \
      jq -r '.[] | select(.StringId | startswith("0x0100580DB2292968A34EA3748511017A6DD2")) | .StringId')
    if [ -z "$contentType" ]; then
      o365 spo list contenttype add --webUrl $portalUrl --listTitle PnP-PortalFooter-Links \
        --contentTypeId 0x0100580DB2292968A34EA3748511017A6DD2 >/dev/null
      success 'DONE'
    else
      warning 'EXISTS'
    fi
    sub '    - Configuring All items view...'
    o365 spo list view set --webUrl $portalUrl --listTitle PnP-PortalFooter-Links \
      --viewTitle 'All Items' \
      --ListViewXml '`<Query><GroupBy Collapse="TRUE" GroupLimit="30"><FieldRef Name="PnPPortalLinkGroup" />  </GroupBy><OrderBy><FieldRef Name="ID" /></OrderBy></Query><ViewFields><FieldRef Name="LinkTitle" /><FieldRef Name="PnPPortalLinkGroup" /><FieldRef Name="PnPPortalLinkUrl" /></ViewFields><RowLimit Paged="TRUE">30</RowLimit><Aggregations Value="Off" /><JSLink>clienttemplates.js</JSLink>`'
    success 'DONE'
    sub '    - List items...\n'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'Find My Customers' \
      --PnPPortalLinkGroup Applications \
      --PnPPortalLinkUrl 'https://find.customers'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'CRM' \
      --PnPPortalLinkGroup Applications \
      --PnPPortalLinkUrl 'https://company.crm'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'ERP' \
      --PnPPortalLinkGroup Applications \
      --PnPPortalLinkUrl 'https://company.erp'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'Technical Procedures' \
      --PnPPortalLinkGroup Applications \
      --PnPPortalLinkUrl 'https://tech.procs'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'Expense Report Module' \
      --PnPPortalLinkGroup 'Internal Modules' \
      --PnPPortalLinkUrl 'https://expense.report'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'Company Car Replacement' \
      --PnPPortalLinkGroup 'Internal Modules' \
      --PnPPortalLinkUrl 'https://need.new.car'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'Vacation Request Module' \
      --PnPPortalLinkGroup 'Internal Modules' \
      --PnPPortalLinkUrl 'https://need.some.rest'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'CNN' \
      --PnPPortalLinkGroup News \
      --PnPPortalLinkUrl 'https://www.cnn.com/'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'BBC' \
      --PnPPortalLinkGroup News \
      --PnPPortalLinkUrl 'https://www.bbc.co.uk/'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'New York Times' \
      --PnPPortalLinkGroup News \
      --PnPPortalLinkUrl 'https://www.nytimes.com/'
    addOrUpdateListItem $portalUrl PnP-PortalFooter-Links 'Forbes' \
      --PnPPortalLinkGroup News \
      --PnPPortalLinkUrl 'http://www.forbes.com/'
    # /PnP-PortalFooter-Links

    checkPoint=260
  fi
fi

if (( $checkPoint < 270 )); then
  # files
  sub '- Provisioning assets...\n'
  files=(
    'Commercial16_smallmeeting_02.jpg'
    'MSSurface_Pro4_SMB_Seattle_0578.jpg'
    'WCO18_hallwayWalk_004.jpg'
    'WCO18_ITHelp_004.jpg'
    'WCO18_whiteboard_002.jpg'
    'Win17_15021_00_N9.jpg'
    'contoso_sitelogo.png'
    'hero.jpg'
    'logo_hr.png'
    'logo_marketing.png'
    'meeting-rooms.jpg'
    'modernOffice_002.jpg'
    'modernOffice_007.jpg'
    'modernOffice_011.jpg'
    'page-faq.jpg'
    'page-feedback.jpg'
    'page-financial-results.jpg'
    'page-hr.jpg'
    'page-my-profile.jpg'
    'page-people-directory.jpg'
    'page-support.jpg'
    'page-training.jpg'
    'page-travel-instructions.jpg'
    'work-life-balance.png'
    'working-methods.jpg'
  )
  for file in "${files[@]}"; do
    sub "  - $file..."
    o365 spo file add --webUrl $portalUrl --folder SiteAssets --path "./resources/images/$file"
    success 'DONE'
  done
  sub "  - contoso_report.pptx..."
  o365 spo file add --webUrl $portalUrl --folder 'Shared Documents' --path "./resources/documents/contoso_report.pptx"
  success 'DONE'
  # /files

  checkPoint=270
fi

if (( $checkPoint < 280 )); then
  # IDs required for provisioning web parts
  sub '- Retrieving ID for file hero.jpg...'
  heroJpgId=$(o365 spo file get --webUrl $portalUrl --url /sites/$(echo $prefix)portal/SiteAssets/hero.jpg --output json | jq -r '.UniqueId')
  success 'DONE'
  sub '- Retrieving ID for file modernOffice_007.jpg...'
  modernOffice_007JpgId=$(o365 spo file get --webUrl $portalUrl --url /sites/$(echo $prefix)portal/SiteAssets/modernOffice_007.jpg --output json | jq -r '.UniqueId')
  success 'DONE'
  sub '- Retrieving ID for file modernOffice_011.jpg...'
  modernOffice_011JpgId=$(o365 spo file get --webUrl $portalUrl --url /sites/$(echo $prefix)portal/SiteAssets/modernOffice_011.jpg --output json | jq -r '.UniqueId')
  success 'DONE'
  sub '- Retrieving ID for file WCO18_hallwayWalk_004.jpg...'
  WCO18_hallwayWalk_004JpgId=$(o365 spo file get --webUrl $portalUrl --url /sites/$(echo $prefix)portal/SiteAssets/WCO18_hallwayWalk_004.jpg --output json | jq -r '.UniqueId')
  success 'DONE'
  sub '- Retrieving ID for file contoso_report.pptx...'
  contoso_reportPptxId=$(o365 spo file get --webUrl $portalUrl --url "/sites/$(echo $prefix)portal/Shared Documents/contoso_report.pptx" --output json | jq -r '.UniqueId')
  success 'DONE'

  sub '- Provisioning pages...\n'
  sub '  - Creating pages...\n'
  pages=(
    'Name:"home.aspx" Title:"Home" Layout:"Home" PromoteAsNewsArticle:"false"'
    'Name:"About-Us.aspx" Title:"About Us" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"HR.aspx" Title:"HR" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"People-Directory.aspx" Title:"People Directory" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"My-Profile.aspx" Title:"My Profile" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"Travel-Instructions.aspx" Title:"Travel Instructions" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"Financial-Results.aspx" Title:"Financial Results" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"FAQ.aspx" Title:"FAQ" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"Training.aspx" Title:"Training" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"Support.aspx" Title:"Support" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"Feedback.aspx" Title:"Feedback" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"Personal.aspx" Title:"Personal" Layout:"Article" PromoteAsNewsArticle:"false"'
    'Name:"Meeting-on-Marketing-In-Non-English-Speaking-Markets-This-Friday.aspx" Title:"Meeting on Marketing In Non-English-Speaking Markets This Friday" Layout:"Article" PromoteAsNewsArticle:"true"'
    'Name:"Marketing-Lunch.aspx" Title:"Marketing lunch" Layout:"Article" PromoteAsNewsArticle:"true"'
    'Name:"New-International-Marketing-Initiatives.aspx" Title:"New International Marketing Initiatives" Layout:"Article" PromoteAsNewsArticle:"true"'
    'Name:"New-Portal.aspx" Title:"New intranet portal" Layout:"Article" PromoteAsNewsArticle:"true"'
  )

  for pageInfo in "${pages[@]}"; do
    pageName=$(getPropertyValue "$pageInfo" "Name")
    pageTitle=$(getPropertyValue "$pageInfo" "Title")
    layout=$(getPropertyValue "$pageInfo" "Layout")
    promoteAsNews=$(getPropertyValue "$pageInfo" "PromoteAsNewsArticle")
    promote=$(if $promoteAsNews = 'true'; then echo "--promoteAs NewsPage"; else echo ""; fi)
    sub "    - $pageName..."
    page=$(o365 spo page get --webUrl $portalUrl --name $pageName --output json || true)
    if ! isError "$page"; then
      warning 'EXISTS'
      warningMsg "    - Removing $pageName..."
      o365 spo page remove --webUrl $portalUrl --name $pageName --confirm
      success 'DONE'
      sub "    - Creating $pageName..."
    fi
    o365 spo page add --webUrl $portalUrl --name $pageName --title "$pageTitle" \
      --layoutType $layout $promote --publish
    success 'DONE'
  done

  sub '  - Configuring headers...\n'
  pages=(
    'Name:"About-Us.aspx" Image:"hero.jpg" X:"42.3837520042758" Y:"56.4285714285714"'
    'Name:"HR.aspx" Image:"page-hr.jpg" X:"44.5216461785142" Y:"53.9285714285714"'
    'Name:"People-Directory.aspx" Image:"page-people-directory.jpg" X:"50.1336183858899" Y:"30"'
    'Name:"My-Profile.aspx" Image:"page-my-profile.jpg" X:"46.4457509353287" Y:"38.2142857142857"'
    'Name:"Travel-Instructions.aspx" Image:"page-travel-instructions.jpg" X:"51.6835916622127" Y:"67.8571428571429"'
    'Name:"Financial-Results.aspx" Image:"page-financial-results.jpg" X:"50.0801710315339" Y:"75.7142857142857"'
    'Name:"FAQ.aspx" Image:"page-faq.jpg" X:"45.6440406199893" Y:"64.2857142857143"'
    'Name:"Training.aspx" Image:"page-training.jpg" X:"51.4163548904329" Y:"15.3571428571429"'
    'Name:"Support.aspx" Image:"page-support.jpg" X:" " Y:" "'
    'Name:"Feedback.aspx" Image:"page-feedback.jpg" X:"48.9043292357028" Y:"33.2142857142857"'
    'Name:"Personal.aspx" Image:"modernOffice_002.jpg" X:"22.0604099244876" Y:"49.6428571428571"'
    'Name:"Meeting-on-Marketing-In-Non-English-Speaking-Markets-This-Friday.aspx" Image:"MSSurface_Pro4_SMB_Seattle_0578.jpg" X:"44.4279786603438" Y:"28.9285714285714"'
    'Name:"Marketing-Lunch.aspx" Image:"Commercial16_smallmeeting_02.jpg" X:"43.1238885595732" Y:"28.5714285714286"'
    'Name:"New-International-Marketing-Initiatives.aspx" Image:"Win17_15021_00_N9.jpg" X:"35.8006487761722" Y:"55.3571428571429"'
    'Name:"New-Portal.aspx" Image:"WCO18_ITHelp_004.jpg" X:"38.0713653789443" Y:"66.7857142857143"'
  )

  for pageInfo in "${pages[@]}"; do
    pageName=$(getPropertyValue "$pageInfo" "Name")
    image=$(getPropertyValue "$pageInfo" "Image")
    x=$(getPropertyValue "$pageInfo" "X")
    y=$(getPropertyValue "$pageInfo" "Y")
    sub "    - $pageName..."
    o365 spo page header set --webUrl $portalUrl --pageName $pageName \
      --type Custom \
      --imageUrl "/sites/$(echo $prefix)portal/SiteAssets/$image" \
      --translateX $x \
      --translateY $y
    success 'DONE'
  done

  sub '  - Provisioning page contents...\n'
  sub '    - home.aspx...\n'
  pageName=home.aspx
  sub '      - Sections...\n'
  sections=(
    'Template:"TwoColumnLeft" Order:"1"'
    'Template:"TwoColumnLeft" Order:"2"'
    'Template:"TwoColumnLeft" Order:"3"'
    'Template:"ThreeColumn" Order:"4"'
    'Template:"ThreeColumn" Order:"5"'
  )
  for section in "${sections[@]}"; do
    template=$(getPropertyValue "$section" "Template")
    order=$(getPropertyValue "$section" "Order")
    sub "        - $order..."
    o365 spo page section add --webUrl $portalUrl --name $pageName \
      --sectionTemplate $template --order $order
    success 'DONE'
  done
  sub '      - Web parts...\n'
  sub '        - Hero...'
  webPartData='{ "dataVersion": "1.3", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{"content[0].title":"About Us","content[1].title":"Human Resources","content[2].title":"Mission","content[3].title":"Projects\n","content[4].title":"Organization","content[0].alternateText":"Company default image","content[1].alternateText":"","content[2].alternateText":"","content[3].alternateText":"","content[4].alternateText":"","content[0].callToActionText":"Learn more","content[1].callToActionText":"Learn more","content[2].callToActionText":"Learn more","content[3].callToActionText":"Learn more","content[4].callToActionText":"Learn more"},"imageSources":{"content[0].image.url":"{site}/SiteAssets/hero.jpg","content[2].image.url":"{site}/SiteAssets/modernOffice_007.jpg","content[3].image.url":"{site}/SiteAssets/modernOffice_011.jpg","content[0].previewImage.url":"{site}/SitePages/About-Us.aspx","content[1].previewImage.url":"https://www.bing.com","content[2].previewImage.url":"{site}/SiteAssets/WCO18_hallwayWalk_004.jpg","content[3].previewImage.url":"{site}/SiteAssets/modernOffice_011.jpg","content[4].previewImage.url":"https://www.bing.com"},"links":{"content[0].link":"{site}/SitePages/About-Us.aspx","content[1].link":"https://www.bing.com","content[2].link":"{site}/SiteAssets/WCO18_hallwayWalk_004.jpg","content[3].link":"{site}/SiteAssets/modernOffice_011.jpg","content[4].link":"https://www.bing.com","content[0].callToActionLink":"{site}/SiteAssets/modernOffice_002.jpg"},"componentDependencies":{"heroLayoutComponentId":"9586b262-54de-4b27-9eb9-34c671400c33","carouselLayoutComponentId":"8ac0c53c-e8d0-4e3e-87d0-7449eb0d4027"},"customMetadata":{"heroLayoutComponentId":{"minCanvasWidth":640},"carouselLayoutComponentId":{"maxCanvasWidth":639}}}, "properties": {"heroLayoutThreshold":640,"carouselLayoutMaxWidth":639,"isFullWidth":true,"layoutCategory":1,"layout":5,"content":[{"id":"9eff894c-f593-4427-a706-413b1cef88c1","type":"Web Page","color":4,"image":{"zoomRatio":1,"siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{listid:Site Assets}","id":"{heroJpgId}"},"description":"","showDescription":false,"showTitle":true,"imageDisplayOption":3,"isDefaultImage":false,"showCallToAction":true,"isDefaultImageLoaded":false,"isCustomImageLoaded":true,"showFeatureText":false,"previewImage":{"zoomRatio":1,"siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{listid:Site Pages}"}},{"id":"4c8a3a11-58eb-45c5-9c68-e7171a4511c5","type":"UrlLink","color":5,"description":"","showDescription":false,"showTitle":true,"imageDisplayOption":2,"isDefaultImage":false,"showCallToAction":false,"isDefaultImageLoaded":false,"isCustomImageLoaded":false,"showFeatureText":false,"previewImage":{"minCanvasWidth":32767}},{"id":"ffff73bd-e508-41ab-990f-ebe6c23939ba","type":"Image","color":4,"image":{"siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{listid:Site Assets}","id":"{modernOffice_007JpgId}","minCanvasWidth":32767},"description":"","showDescription":false,"showTitle":true,"imageDisplayOption":1,"isDefaultImage":false,"showCallToAction":false,"isDefaultImageLoaded":true,"isCustomImageLoaded":false,"showFeatureText":false,"previewImage":{"siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{{listid:Site Assets}}","id":"{WCO18_hallwayWalk_004JpgId}","widthFactor":0.25,"minCanvasWidth":640}},{"id":"5972b83b-7497-4d20-bd1a-8c50c96ac2fc","type":"Image","color":4,"image":{"siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{listid:Site Assets}","id":"{modernOffice_011JpgId}","minCanvasWidth":32767},"description":"","showDescription":false,"showTitle":true,"imageDisplayOption":1,"isDefaultImage":false,"showCallToAction":false,"isDefaultImageLoaded":true,"isCustomImageLoaded":false,"showFeatureText":false,"previewImage":{"siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{listid:Site Assets}","id":"{modernOffice_011JpgId}","widthFactor":0.25,"minCanvasWidth":32767}},{"id":"f743acc1-186b-4699-be7a-6a9773ae08f7","type":"UrlLink","color":5,"description":"","showDescription":false,"showTitle":true,"imageDisplayOption":2,"isDefaultImage":false,"showCallToAction":false,"isDefaultImageLoaded":false,"isCustomImageLoaded":false,"showFeatureText":false,"previewImage":{"minCanvasWidth":32767}}]}}'
  webPartData=$(echo "${webPartData//\{sitecollectionid\}/$siteId}")
  webPartData=$(echo "${webPartData//\{siteid\}/$webId}")
  webPartData=$(echo "${webPartData//\{listid:Site Assets\}/$siteAssetsListId}")
  webPartData=$(echo "${webPartData//\{listid:Site Pages\}/$sitePagesListId}")
  webPartData=$(echo "${webPartData//\{site\}/$portalUrl}")
  webPartData=$(echo "${webPartData//\{heroJpgId\}/$heroJpgId}")
  webPartData=$(echo "${webPartData//\{modernOffice_007JpgId\}/$modernOffice_007JpgId}")
  webPartData=$(echo "${webPartData//\{modernOffice_011JpgId\}/$modernOffice_011JpgId}")
  webPartData=$(echo "${webPartData//\{WCO18_hallwayWalk_004JpgId\}/$WCO18_hallwayWalk_004JpgId}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --standardWebPart Hero \
    --section 1 --column 1 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - Tiles...'
  webPartData='{ "dataVersion": "1.0", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}}, "properties": {"title":"","listUrl":"{hosturl}{site}/Lists/PnPTiles","collectionData":[{"title":"Employee Directory","description":"Enterprise Phone book","url":"{site}/SitePages/People-Directory.aspx","icon":"Group","target":""},{"title":"HR","description":"Human Resources","url":"{site}/SitePages/HR.aspx","icon":"managerselfservice","target":""},{"title":"My Profile","description":"My Profile","url":"{site}/SitePages/My-Profile.aspx","icon":"d365talenthrcore","target":""},{"title":"Travel Instructions","description":"Traveling?","url":"{site}/SitePages/Travel-Instructions.aspx","icon":"airplane","target":""},{"title":"Financial Results","description":"Company Results","url":"{site}/SitePages/Financial-Results.aspx","icon":"stackedlinechart","target":""},{"title":"FAQ","description":"Frequently Asked Questions","url":"{site}/SitePages/FAQ.aspx","icon":"searchissue","target":""},{"title":"Training","description":"Training materials","url":"{site}/SitePages/Training.aspx","icon":"d365talentlearn","target":""},{"title":"Support","description":"Support","url":"{site}/SitePages/Support.aspx","icon":"headset","target":""},{"title":"Feedback","description":"Provide feedback on portal","url":"{site}/SitePages/Feedback.aspx","icon":"ChatInviteFriend","target":""}]}}'
  webPartData=$(echo "${webPartData//\{site\}/$portalUrl}")
  webPartData=$(echo "${webPartData//\{hosturl\}/}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 26cb4af3-7f48-4737-b82a-4e24167c2d07 \
    --section 1 --column 2 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - NewsReel...'
  webPartData='{ "dataVersion": "1.5", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{"title":"Company News"},"imageSources":{},"links":{"baseUrl":"{hosturl}{site}"},"componentDependencies":{"layoutComponentId":"a2752e70-c076-41bf-a42e-1d955b449fbc"}}, "properties": {"layoutId":"FeaturedNews","filters":[{"filterType":1,"value":"","values":[]}],"newsDataSourceProp":1,"dataProviderId":"viewCounts","newsSiteList":[],"renderItemsSliderValue":3,"webId":"{siteid}","siteId":"{sitecollectionid}","templateId":"FeaturedNews","propsLastEdited":"2018-07-09T22:55:36.594Z","showChrome":true,"prefetchCount":4,"compactMode":false}}'
  webPartData=$(echo "${webPartData//\{site\}/$portalUrl}")
  webPartData=$(echo "${webPartData//\{hosturl\}/}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --standardWebPart NewsReel \
    --section 2 --column 1 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - Events...'
  webPartData='{ "dataVersion": "1.2", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{"title":"Company Events"},"imageSources":{},"links":{"baseUrl":"{hosturl}{site}"},"componentDependencies":{"layoutComponentId":"0447e11d-bed9-4898-b600-8dbcd95e9cc2"}}, "properties": {"selectedListId":"{listid:Events}","selectedCategory":"","dateRangeOption":0,"startDate":"","endDate":"","isOnSeeAllPage":false,"layoutId":"Flex","dataProviderId":"Event","webId":"{siteid}","siteId":"{sitecollectionid}","layout":"Filmstrip","dataSource":7,"sites":[],"maxItemsPerPage":20}}'
  webPartData=$(echo "${webPartData//\{sitecollectionid\}/$siteId}")
  webPartData=$(echo "${webPartData//\{siteid\}/$webId}")
  webPartData=$(echo "${webPartData//\{site\}/$portalUrl}")
  webPartData=$(echo "${webPartData//\{hosturl\}/}")
  webPartData=$(echo "${webPartData//\{listid:Events\}/$eventsListId}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --standardWebPart Events \
    --section 2 --column 2 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - DocumentEmbed...'
  webPartData='{ "dataVersion": "1.2", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{"annotation":"Latest quarterly company presentation","title":"PowerPoint Presentation"},"imageSources":{},"links":{"serverRelativeUrl":"{site}/Shared Documents/Contoso_Report.pptx","wopiurl":"{hosturl}{site}/_layouts/15/WopiFrame.aspx?sourcedoc={{fileuniqueid:/shared documents/contoso_report.pptx}}&amp;action=interactivepreview"}}, "properties": {"authorName":"Vesa Juvonen","chartitem":"","endrange":"","excelSettingsType":"","file":"{hosturl}{site}/Shared Documents/Contoso_Report.pptx","listId":"{listid:Documents}","modifiedAt":"2018-05-14T13:40:02+02:00","photoUrl":"/_layouts/15/userphoto.aspx?size=S&amp;accountname=vesaj%40officedevpnp.onmicrosoft.com","rangeitem":"","siteId":"{sitecollectionid}","startPage":1,"startrange":"","tableitem":"","uniqueId":"{fileuniqueid:/shared documents/contoso_report.pptx}","wdallowinteractivity":true,"wdhidegridlines":true,"wdhideheaders":true,"webId":"{siteid}"}}'
  webPartData=$(echo "${webPartData//\{sitecollectionid\}/$siteId}")
  webPartData=$(echo "${webPartData//\{siteid\}/$webId}")
  webPartData=$(echo "${webPartData//\{site\}/$portalUrl}")
  webPartData=$(echo "${webPartData//\{hosturl\}/}")
  webPartData=$(echo "${webPartData//\{fileuniqueid:\/shared documents\/contoso_report.pptx\}/$contoso_reportPptxId}")
  webPartData=$(echo "${webPartData//\{listid:Documents\}/$documentsListId}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --standardWebPart DocumentEmbed \
    --section 3 --column 1 --order 2 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - Weather...'
  webPartData='{ "dataVersion": "1.0", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}}, "properties": {"temperatureUnit":"F","location":{"latitude":28.538,"longitude":-81.377,"name":"{parameter:weatherCity}"}}}'
  webPartData=$(echo "${webPartData//\{parameter:weatherCity\}/$weatherCity}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823 \
    --section 3 --column 2 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - Stocks...'
  webPartData='{ "dataVersion": "1.0", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}}, "properties": {"stockSymbol":"{parameter:StockSymbol}","autoRefresh":false,"demo":true}}'
  webPartData=$(echo "${webPartData//\{parameter:StockSymbol\}/$stockSymbol}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 50256fc2-a28f-4544-900e-32724d32bc7f \
    --section 3 --column 2 --order 2 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - Image #1...'
  webPartData='{ "dataVersion": "1.8", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{"captionText":""},"imageSources":{"imageSource":"{site}/SiteAssets/working-methods.jpg"},"links":{}}, "properties": {"imageSourceType":2,"altText":"a person sitting at a table in a room","overlayText":"Working methods","fileName":"87461-GettyImages-867431800.jpg","siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{listid:Site Assets}","uniqueId":"dd4d43e7-2ac8-4d43-a1bd-dcb05672a1be","imgWidth":1500,"imgHeight":1000,"fixAspectRatio":false,"isOverlayTextEnabled":true}}'
  webPartData=$(echo "${webPartData//\{sitecollectionid\}/$siteId}")
  webPartData=$(echo "${webPartData//\{siteid\}/$webId}")
  webPartData=$(echo "${webPartData//\{listid:Site Assets\}/$siteAssetsListId}")
  webPartData=$(echo "${webPartData//\{site\}/$portalUrl}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --standardWebPart Image \
    --section 4 --column 1 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - Image #2...'
  webPartData='{ "dataVersion": "1.8", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{"captionText":""},"imageSources":{"imageSource":"{site}/SiteAssets/meeting-rooms.jpg"},"links":{}}, "properties": {"imageSourceType":2,"altText":"a flat screen tv sitting in a room","overlayText":"Meeting Rooms","fileName":"4335-WCO18_slidingDoor_001.jpg","siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{listid:Site Assets}","uniqueId":"2d8b9627-e91d-412d-aacc-3ec19f59bc1d","imgWidth":1500,"imgHeight":1000,"fixAspectRatio":false,"isOverlayTextEnabled":true}}'
  webPartData=$(echo "${webPartData//\{sitecollectionid\}/$siteId}")
  webPartData=$(echo "${webPartData//\{siteid\}/$webId}")
  webPartData=$(echo "${webPartData//\{listid:Site Assets\}/$siteAssetsListId}")
  webPartData=$(echo "${webPartData//\{site\}/$portalUrl}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --standardWebPart Image \
    --section 4 --column 2 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - Image #3...'
  webPartData='{ "dataVersion": "1.8", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{"captionText":""},"imageSources":{"imageSource":"{site}/SiteAssets/work-life-balance.png"},"links":{}}, "properties": {"imageSourceType":2,"altText":"a group of people on a beach","overlayText":"Work life balance","fileName":"48146-OFF12_Justice_01.png","siteId":"{sitecollectionid}","webId":"{siteid}","listId":"{listid:Site Assets}","uniqueId":"67664b85-067d-4be9-a7d7-89b2e804d09f","imgWidth":650,"imgHeight":433,"fixAspectRatio":false,"isOverlayTextEnabled":true}}'
  webPartData=$(echo "${webPartData//\{sitecollectionid\}/$siteId}")
  webPartData=$(echo "${webPartData//\{siteid\}/$webId}")
  webPartData=$(echo "${webPartData//\{listid:Site Assets\}/$siteAssetsListId}")
  webPartData=$(echo "${webPartData//\{site\}/$portalUrl}")
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --standardWebPart Image \
    --section 4 --column 3 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - World clock Singapore...'
  webPartData='{ "dataVersion": "1.0", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}}, "properties": {"description":"Singapore","timeZoneOffset":103}}'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 4f87b698-f910-451f-b4ea-7848a472af0f \
    --section 5 --column 1 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - World clock London...'
  webPartData='{ "dataVersion": "1.0", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}}, "properties": {"description":"London","timeZoneOffset":48}}'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 4f87b698-f910-451f-b4ea-7848a472af0f \
    --section 5 --column 2 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'
  sub '        - World clock Seattle...'
  webPartData='{ "dataVersion": "1.0", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}}, "properties": {"description":"Seattle","timeZoneOffset":10}}'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 4f87b698-f910-451f-b4ea-7848a472af0f \
    --section 5 --column 3 --order 1 \
    --webPartData '`'"$webPartData"'`'
  success 'DONE'

  sub '    - personal.aspx...\n'
  pageName=personal.aspx
  sub '      - Sections...\n'
  sections=(
    'Template:"ThreeColumn" Order:"1"'
    'Template:"OneColumn" Order:"2"'
  )
  for section in "${sections[@]}"; do
    template=$(getPropertyValue "$section" "Template")
    order=$(getPropertyValue "$section" "Order")
    sub "        - $order..."
    o365 spo page section add --webUrl $portalUrl --name $pageName \
      --sectionTemplate $template --order $order
    success 'DONE'
  done
  sub '      - Web parts...\n'
  sub '        - Personal calendar...'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId bbcd646e-6b86-4480-a68a-850c98f94519 \
    --section 1 --column 1 --order 1
  success 'DONE'
  sub '        - Personal contacts...'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId a95742b9-5e67-4a11-8e51-7a76812a9d60 \
    --section 1 --column 1 --order 2
  success 'DONE'
  sub '        - Personal tasks...'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId ee972022-9937-4b69-9c25-eaf77003f4f9 \
    --section 1 --column 2 --order 1
  success 'DONE'
  sub '        - Personal email...'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId ac41debd-d92c-4de0-9d67-e7bc191030ee \
    --section 1 --column 2 --order 2
  success 'DONE'
  sub '        - Followed sites...'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 92d5cc6c-fae5-4e91-9583-ab33950f5a8d \
    --section 1 --column 3 --order 1
  success 'DONE'
  sub '        - Recently used documents...'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 596e3e97-f3ca-4cf9-9b8e-cb852f063356 \
    --section 1 --column 3 --order 2
  success 'DONE'
  sub '        - Recently visited sites...'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId 277282d0-debb-42d6-b4b3-30ec543964fd \
    --section 1 --column 3 --order 3
  success 'DONE'
  sub '        - People directory...'
  o365 spo page clientsidewebpart add --webUrl $portalUrl --pageName $pageName \
    --webPartId d9d6f014-3fd9-4fb9-af54-044c7c43e081 \
    --section 2 --column 1 --order 1
  success 'DONE'

  sub '    - Meeting-on-Marketing-In-Non-English-Speaking-Markets-This-Friday.aspx...\n'
  pageName=Meeting-on-Marketing-In-Non-English-Speaking-Markets-This-Friday.aspx
  sub '      - Section...'
  o365 spo page section add --webUrl $portalUrl --name $pageName \
      --sectionTemplate OneColumn --order 1
  success 'DONE'
  sub '      - Text...'
  o365 spo page text add --webUrl $portalUrl --pageName $pageName \
    --text 'Please attend this department-wide meeting next Thursday at 4:00PM in conference room seven. We will be discussing tactics on how to effectively create marketing campaigns in our new international markets.'
  success 'DONE'

  sub '    - Marketing-Lunch...\n'
  pageName=Marketing-Lunch.aspx
  sub '      - Section...'
  o365 spo page section add --webUrl $portalUrl --name $pageName \
      --sectionTemplate OneColumn --order 1
  success 'DONE'
  sub '      - Text...'
  o365 spo page text add --webUrl $portalUrl --pageName $pageName \
    --text 'There is a lunch for the Marketing team Next Tuesday at 12:30PM. All Marketing team members should attend, as we will be talking about some of the new projects that we hope to put through for the holiday season.'
  success 'DONE'

  sub '    - New-International-Marketing-Initiatives...\n'
  pageName=New-International-Marketing-Initiatives.aspx
  sub '      - Section...'
  o365 spo page section add --webUrl $portalUrl --name $pageName \
      --sectionTemplate OneColumn --order 1
  success 'DONE'
  sub '      - Text...'
  o365 spo page text add --webUrl $portalUrl --pageName $pageName \
    --text 'We will be releasing a new international marketing campaign in the coming weeks. Look for more details here.​'
  success 'DONE'

  sub '    - New-Portal...\n'
  pageName=New-Portal.aspx
  sub '      - Section...'
  o365 spo page section add --webUrl $portalUrl --name $pageName \
      --sectionTemplate OneColumn --order 1
  success 'DONE'
  sub '      - Text...'
  o365 spo page text add --webUrl $portalUrl --pageName $pageName \
    --text 'We are happy to announce availability of this new Intranet portal. Please do give us feedback!'
  success 'DONE'

  checkPoint=280
fi

if (( $checkPoint < 290 )); then
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

  checkPoint=290
fi

if (( $checkPoint < 295 )); then
  setupPortalExtensions $portalUrl

  checkPoint=295
fi

if (( $checkPoint < 500 )); then
  success 'DONE'
  echo
fi

if (( $checkPoint < 300 )); then
  checkPoint=300
fi