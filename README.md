# Setup bash script

Setup bash script for the [PnP Starter Kit](https://github.com/SharePoint/sp-starter-kit)

## Prerequisites

- Install
  - [Office 365 CLI](https://aka.ms/o365cli) latest beta (`npm i -g @pnp/office365-cli@next`)
  - [jq](https://stedolan.github.io/jq/)
- Configure
  - Set the user who will run the setup script as a taxonomy term store admin
- Execute
  - `o365 spo connect [tenant admin]`, eg. `o365 spo connect https://contoso-admin.sharepoint.com`
  - `o365 graph connect`
  - `chmod +x ./setup.sh`

## Setup

Execute the setup script by in the command line running: `./setup.sh --tenantUrl https://contoso.sharepoint.com --prefix pnp_`.

Following are the options you can pass to the script:

argument|description|required|default value|example value
--------|-----------|--------|-------------|-------------
`-t, --tenantUrl`|URL of the SharePoint tenant where the Starter Kit should be provisioned|yes|`undefined`|`https://contoso.sharepoint.com`
`-p, --prefix`|Prefix to use when provisioning sites to avoid conflicts with existing sites|no|`(empty string)`|`starterkit`
`--skipSolutionDeployment`|Don't deploy the solution package|no|`false`|`true`
`--skipSiteCreation`|Don't create sites|no|`false`|`true`
`--stockAPIKey`|API key to use with the Alpha Vantage API to retrieve stock information|no|`(empty string)`|`12345`
`-c, --company`|Name of the organization to use when provisioning Starter Kit|no|`Contoso`|`Contoso`
`-w, --weatherCity`|City for which to display weather in the weather web part|no|`Seattle`|`Seattle`
`--stockSymbol`|Symbol of the stock to display in the stock web part|no|`MSFT`|`MSFT`

## Remove

To remove the Starter Kit, execute the uninstall script by in the command line running: `./remove.sh --tenantUrl https://contoso.sharepoint.com --prefix pnp_`.

Following are the options you can pass to the script:

argument|description|required|default value|example value
--------|-----------|--------|-------------|-------------
`-t, --tenantUrl`|URL of the SharePoint tenant where the Starter Kit should be provisioned|yes|`undefined`|`https://contoso.sharepoint.com`
`-p, --prefix`|Prefix to use when provisioning sites to avoid conflicts with existing sites|no|`(empty string)`|`starterkit`
`-c, --company`|Name of the organization to use when provisioning Starter Kit|no|`Contoso`|`Contoso`