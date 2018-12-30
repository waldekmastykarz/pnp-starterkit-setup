siteUrl=$marketingUrl
alias=$(echo $prefix)marketing
msg "Provisioning marketing site at $siteUrl..."
if (( $checkPoint >= 410 )); then
  msg "\n"
fi
if (( $checkPoint < 410 )); then
  site=$(o365 spo site get --url $siteUrl --output json || true)
  if $(isError "$site"); then
    o365 spo site add --type TeamSite --url $siteUrl --title Marketing --alias $alias --lcid 1033 >/dev/null
    success 'DONE'
  else
    warning 'EXISTS'
  fi

  checkPoint=410
fi

if (( $checkPoint < 420 )); then
  sub '- Connecting to the hub site...'
  o365 spo hubsite connect --url $siteUrl --hubSiteId $hubsiteId
  success 'DONE'

  checkPoint=420
fi

if (( $checkPoint < 430 )); then
  sub '- Applying theme...'
  o365 spo theme apply --name "$company Marketing" --webUrl $siteUrl >/dev/null
  success 'DONE'

  checkPoint=430
fi

if (( $checkPoint < 440 )); then
  setupCollabExtensions $siteUrl

  checkPoint=440
fi

if (( $checkPoint < 450 )); then
  sub '- Setting logo...'
  groupId=$(o365 graph o365group list --mailNickname $alias -o json | jq -r '.[] | select(.mailNickname == "'"$alias"'") | .id')
  if [ -z "$groupId" ]; then
    error 'ERROR'
    error "Office 365 Group '$alias' not found"
    exit 1
  fi
  o365 graph o365group set --id $groupId --logoPath ./resources/images/logo_marketing.png
  success 'DONE'

  checkPoint=450
fi

success 'DONE'
echo

checkPoint=500