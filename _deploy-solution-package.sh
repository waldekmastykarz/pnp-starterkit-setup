if [ ! $skipSolutionDeployment = true ]; then
  echo "Deploying solution..."

  app=$(o365 spo app get --name "sharepoint-portal-showcase.sppkg" --output json || true)
  if ! isError "$app"; then
    warning "Solution package already exists. Removing..."
    appId=$(echo $app | jq -r '.ID')
    o365 spo app remove --id $appId --confirm
    success "DONE"
  fi

  echo "Adding solution package to tenant app catalog..."
  o365 spo app add --filePath ./sharepoint-portal-showcase.sppkg
  success "DONE"

  echo "Deploying solution package..."
  o365 spo app deploy --name sharepoint-portal-showcase.sppkg --skipFeatureDeployment
  success "DONE"
fi