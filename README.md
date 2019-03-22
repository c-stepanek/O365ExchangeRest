# O365ExchangeRest
This is a PowerShell module for Exchange in Office365 leveraging the Exchange.RestServices API by Ivan Franjic.

Once you build the solution you will need some setup to get the module working with your tenant.
You can refer to Ivan's documentation here: https://github.com/ivfranji/Exchange.RestServices/wiki

# Module Config
Config File: `O365ExchangeRest.dll.config`

Use `New-O365ExchangeServiceCertificate` to create a self-signed certificate and generate the JSON blob for the Azure application manifest.

Use `Get-O365ExchangeRestAppConfig` to query the config file.

Use `Set-O365ExchangeRestAppConfig` to update the config file.

# Notes
If you have made changes to the config file after the module has been loaded, you will have to reload the module to pickup the changes.
