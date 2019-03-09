# O365ExchangeRest
This is a PowerShell module for Exchange in Office365 leveraging the Exchange.RestServices API by Ivan Franjic.

Once you build the solution you will need some setup to get the module working with your tenant.
You can refer to Ivan's documentation here: https://github.com/ivfranji/Exchange.RestServices/wiki

# Module Config
Config File: `O365ExchangeRest.dll.config`

I have added `New-O365ExchangeServiceCertificate` to the module which will create the self-signed certificate and update the config file with the certificate thumbprint.
You will need to manually add your Office365 TenantId and Azure ApplicationId.

# Notes
If you have made changes to the config file after the module has been loaded, you will have to reload the module to pickup the changes.
