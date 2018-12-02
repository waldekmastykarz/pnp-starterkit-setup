### term group
msg "Provisioning term group..."
result=$(o365 spo term group get --name PnPTermSets --output json || true)
if $(isError "$result"); then
  result=$(o365 spo term group add --name PnPTermSets --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb --output json || true)
  if $(isError "$result"); then
    error 'ERROR'
    errorMessage "$result"
    exit 1
  fi
  success 'DONE'
else
  warning 'EXISTS'
fi
echo
### /term group

### term sets
msg 'Provisioning term sets...\n'
termSets=(
  'ID:"7a167c47-2b37-41d0-94d0-e962c1a4f2ed" Title:"PnP-CollabFooter-SharedLinks"',
  'ID:"1479e26c-1380-41a8-9183-72bc5a9651bb" Title:"PnP-Organizations"'
)
for termSet in "${termSets[@]}"; do
  termSetId=$(getPropertyValue "$termSet" "ID")
  termSetTitle=$(getPropertyValue "$termSet" "Title")
  sub "- $termSetTitle..."
  result=$(o365 spo term set get --id $termSetId --termGroupName PnPTermSets --output json || true)
  if $(isError "$result"); then
    result=$(o365 spo term set add --name "$termSetTitle" --id $termSetId --customProperties '`{"_Sys_Nav_IsNavigationTermSet":"True"}`' --termGroupName PnPTermSets --output json || true)
    if $(isError "$result"); then
      error 'ERROR'
      errorMessage "$result"
      exit 1
    fi
    success 'DONE'
  else
    warning 'EXISTS'
  fi
done
success 'DONE'
echo
### /term sets

### PnP-CollabFooter-SharedLinks terms
msg 'Provisioning PnP-CollabFooter-SharedLinks terms...\n'
terms=(
  'ID:"a359ee29-cf72-4235-a4ef-1ed96bf4eaea" Name:"Contacts" LinkUrl:"https://intranet.contoso.com/Contacts" Icon:"EditMail"',
  'ID:"60d165e6-8cb1-4c20-8fad-80067c4ca767" Name:"Legal Policies" LinkUrl:"https://intranet.contoso.com/LegalPolicies" Icon:"Lock"',
  'ID:"da7bfb84-008b-48ff-b61f-bfe40da2602f" Name:"Tools" LinkUrl:" " Icon:"DeveloperTools"'
)
for term in "${terms[@]}"; do
  termId=$(getPropertyValue "$term" "ID")
  termName=$(getPropertyValue "$term" "Name")
  linkUrl=$(getPropertyValue "$term" "LinkUrl")
  icon=$(getPropertyValue "$term" "Icon")
  sub "- $termName..."
  result=$(o365 spo term get --id $termId --output json)
  if [ -z "$result" ]; then
    result=$(o365 spo term add --name "$termName" --id $termId --customProperties '`{"_Sys_Nav_SimpleLinkUrl":"'"$linkUrl"'","PnP-CollabFooter-Icon":"'"$icon"'"}`' --termGroupName PnPTermSets --termSetName PnP-CollabFooter-SharedLinks --output json || true)
    if $(isError "$result"); then
      error 'ERROR'
      errorMessage "$result"
      exit 1
    fi
    success 'DONE'
  else
    warning 'EXISTS'
  fi
done
success 'DONE'
echo
### /PnP-CollabFooter-SharedLinks terms

### PnP-CollabFooter-SharedLinks Tools terms
msg 'Provisioning PnP-CollabFooter-SharedLinks Tools terms...\n'
terms=(
  'ID:"30f4c129-6886-406c-911a-1395250d690f" Name:"Expense Reports" LinkUrl:"https://downloads.contoso.com/ExpenseReports"',
  'ID:"bef5028a-1dff-43b7-b57d-2aefd3f7b814" Name:"Time Reports" LinkUrl:"https://downloads.contoso.com/TimeReports"',
  'ID:"269c23e2-f34e-438a-adf0-22d11e064de5" Name:"WebMail" LinkUrl:"https://mail.office365.com/"'
)
for term in "${terms[@]}"; do
  termId=$(getPropertyValue "$term" "ID")
  termName=$(getPropertyValue "$term" "Name")
  linkUrl=$(getPropertyValue "$term" "LinkUrl")
  sub "- $termName..."
  result=$(o365 spo term get --id $termId --output json)
  if [ -z "$result" ]; then
    result=$(o365 spo term add --name "$termName" --id $termId --customProperties '`{"_Sys_Nav_SimpleLinkUrl":"'"$linkUrl"'"}`' --termGroupName PnPTermSets --parentTermId da7bfb84-008b-48ff-b61f-bfe40da2602f --output json || true)
    if $(isError "$result"); then
      error 'ERROR'
      errorMessage "$result"
      exit 1
    fi
    success 'DONE'
  else
    warning 'EXISTS'
  fi
done
success 'DONE'
echo
### /PnP-CollabFooter-SharedLinks Tools terms

### PnP-Organizations terms
msg 'Provisioning PnP-Organizations terms...\n'
terms=(
  'ID:"02cf219e-8ce9-4e85-ac04-a913a44a5d2b" Name:"HR"',
  'ID:"247543b6-45f2-4232-b9e8-66c5bf53c31e" Name:"IT"',
  'ID:"ffc3608f-1250-4d28-b388-381fad8d4602" Name:"Leadership"',
)
for term in "${terms[@]}"; do
  termId=$(getPropertyValue "$term" "ID")
  termName=$(getPropertyValue "$term" "Name")
  sub "- $termName..."
  result=$(o365 spo term get --id $termId --output json)
  if [ -z "$result" ]; then
    result=$(o365 spo term add --name "$termName" --id $termId --termGroupName PnPTermSets --termSetName PnP-Organizations --output json || true)
    if $(isError "$result"); then
      error 'ERROR'
      errorMessage "$result"
      exit 1
    fi
    success 'DONE'
  else
    warning 'EXISTS'
  fi
done
success 'DONE'
echo
### /PnP-Organizations terms