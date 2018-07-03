if [ ! $skipSolutionDeployment ]; then
  echo "Deploying solution..."

  app=$(o365 spo app get --name "sharepoint-portal-showcase.sppkg" --output json)
  if [ ! -z "$app" ]; then
    echo "Solution package already exists. Removing..."
    o365 spo app remove --name "sharepoint-portal-showcase.sppkg"
    success "DONE"
  fi

  echo "Add solution package to tenant app catalog..."
  o365 spo app add --filePath ./sharepoint-portal-showcase.sppkg
  success "DONE"
fi