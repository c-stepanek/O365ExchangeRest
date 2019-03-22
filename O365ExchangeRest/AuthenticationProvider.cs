//-----------------------------------------------------------------------
// <copyright file="AuthenticationProvider.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Ivan Franjic (ivfranji)</author>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System;
    using System.Net.Http.Headers;
    using System.Security.Cryptography.X509Certificates;
    using Exchange.RestServices;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;

    /// <summary>
    /// Test authentication provider. 
    /// </summary>
    internal class AuthenticationProvider : IAuthorizationTokenProvider
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AuthenticationProvider" /> class.
        /// </summary>
        internal AuthenticationProvider()
        {
            this.ResourceUri = "https://outlook.office365.com";
        }
        
        /// <summary>
        /// Gets the Resource uri.
        /// </summary>
        private string ResourceUri { get; }

        /// <inheritdoc cref="IAuthorizationTokenProvider.GetAuthenticationHeader"/>
        public AuthenticationHeaderValue GetAuthenticationHeader()
        {
            string token = this.GetToken();
            return new AuthenticationHeaderValue("Bearer", token);
        }

        /// <summary>
        /// Retrieve token.
        /// </summary>
        /// <returns>Returns an access token.</returns>
        private string GetToken()
        {
            string applicationId = Utils.GetConfigKeyValue("ApplicationId");
            string certThumbprint = Utils.GetConfigKeyValue("CertThumbprint");
            string tenantId = Utils.GetConfigKeyValue("TenantId");

            string authority = $"https://login.microsoftonline.com/{tenantId}";
            AuthenticationContext context = new AuthenticationContext(authority);

            X509Certificate2 certFromStore = Utils.GetCertificate(
                X509FindType.FindByThumbprint,
                certThumbprint,
                false);

            if (certFromStore == null)
            {
                throw new ArgumentNullException("Certificate", "Make sure you have the proper certificate thumbprint in the module config file.");
            }

            ClientAssertionCertificate cert = new ClientAssertionCertificate(
                applicationId,
                certFromStore);

            AuthenticationResult token;

            try
            {
                token = context.AcquireTokenAsync(this.ResourceUri, cert).Result;
            }
            catch (Exception ex)
            {
                if (ex.InnerException.Message.ToString() == "Keyset does not exist\r\n")
                {
                    throw new UnauthorizedAccessException("You need to be an administrator to read the private key from the local machine store.");
                }
                else
                {
                    throw ex;
                }
            }

            return token.AccessToken;
        }
    }
}
