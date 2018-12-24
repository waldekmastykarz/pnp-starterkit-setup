siteUrl=$hrUrl
alias=$(echo $prefix)hr
msg "Provisioning HR site at $siteUrl..."
if (( $checkPoint >= 310 )); then
  msg "\n"
fi
if (( $checkPoint < 310 )); then
  site=$(o365 spo site get --url $siteUrl --output json || true)
  if $(isError "$site"); then
    o365 spo site add --type TeamSite --url $siteUrl --title "Human Resources" --alias $alias >/dev/null
    success 'DONE'
  else
    warning 'EXISTS'
  fi

  checkPoint=310
fi

if (( $checkPoint < 320 )); then
  sub '- Connecting to the hub site...'
  o365 spo hubsite connect --url $siteUrl --hubSiteId $hubsiteId
  success 'DONE'

  checkPoint=320
fi

if (( $checkPoint < 330 )); then
  sub '- Applying theme...'
  o365 spo theme apply --name "$company HR" --webUrl $siteUrl >/dev/null
  success 'DONE'

  checkPoint=330
fi

if (( $checkPoint < 340 )); then
  setupCollabExtensions $siteUrl

  checkPoint=340
fi

if (( $checkPoint < 350 )); then
  sub '- Setting logo...'
  groupId=$(o365 graph o365group list --mailNickname $alias -o json | jq -r '.[] | select(.mailNickname == "'"$alias"'") | .id')
  if [ -z "$groupId" ]; then
    error 'ERROR'
    error "Office 365 Group '$alias' not found"
    exit 1
  fi
  o365 graph o365group set --id $groupId --logoPath ./resources/images/logo_hr.png
  success 'DONE'

  checkPoint=350
fi

success 'DONE'
echo

checkPoint=400