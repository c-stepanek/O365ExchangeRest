//-----------------------------------------------------------------------
// <copyright file="Utils.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System;
    using System.Linq;
    using System.Reflection;
    using System.Security.Cryptography.X509Certificates;
    using System.Xml.Linq;
    using CERTENROLLLib;

    /// <summary>
    /// Defines the <see cref="Utils"/> class.
    /// </summary>
    public static class Utils
    {
        /// <summary>
        /// Creates a self-signed certificate using the CertEnroll COM library.
        /// </summary>
        /// <param name="subjectName">Subject name for the certificate.</param>
        /// <returns>Returns a certificate</returns>
        public static X509Certificate2 NewSelfSignedCertificate(string subjectName)
        {
            CX500DistinguishedName dn = new CX500DistinguishedName();
            dn.Encode("CN=" + subjectName, X500NameFlags.XCN_CERT_NAME_STR_NONE);

            CX509PrivateKey privateKey = new CX509PrivateKey
            {
                ProviderName = "Microsoft Enhanced RSA and AES Cryptographic Provider",
                MachineContext = true,
                Length = 2048,
                KeySpec = X509KeySpec.XCN_AT_SIGNATURE,
                ExportPolicy = X509PrivateKeyExportFlags.XCN_NCRYPT_ALLOW_PLAINTEXT_EXPORT_FLAG
            };
            privateKey.Create();

            CObjectId hashobj = new CObjectId();
            hashobj.InitializeFromAlgorithmName(
                ObjectIdGroupId.XCN_CRYPT_HASH_ALG_OID_GROUP_ID,
                ObjectIdPublicKeyFlags.XCN_CRYPT_OID_INFO_PUBKEY_ANY,
                AlgorithmFlags.AlgorithmFlagsNone, 
                "SHA1");

            CObjectId clientObjectId = new CObjectId();
            clientObjectId.InitializeFromValue("1.3.6.1.5.5.7.3.2"); // Client Authentication
            CObjectId serverObjectId = new CObjectId();
            serverObjectId.InitializeFromValue("1.3.6.1.5.5.7.3.1"); // Server Authentication
            CObjectIds oidlist = new CObjectIds
            {
                clientObjectId,
                serverObjectId
            };

            CX509ExtensionEnhancedKeyUsage eku = new CX509ExtensionEnhancedKeyUsage();
            eku.InitializeEncode(oidlist);

            // Create the self signing request
            CX509CertificateRequestCertificate cert = new CX509CertificateRequestCertificate();
            cert.InitializeFromPrivateKey(X509CertificateEnrollmentContext.ContextMachine, privateKey, string.Empty);
            cert.Subject = dn;
            cert.Issuer = dn;
            cert.NotBefore = DateTime.Now;
            cert.NotAfter = DateTime.Now.AddYears(1);
            cert.X509Extensions.Add((CX509Extension)eku);
            cert.HashAlgorithm = hashobj;        
            cert.Encode();

            // Do the final enrollment 
            CX509Enrollment enroll = new CX509Enrollment();
            enroll.InitializeFromRequest(cert);
            string csr = enroll.CreateRequest();
            enroll.InstallResponse(
                InstallResponseRestrictionFlags.AllowUntrustedCertificate,
                csr, 
                EncodingType.XCN_CRYPT_STRING_BASE64,
                string.Empty);
            string base64encoded = enroll.CreatePFX(string.Empty, PFXExportOptions.PFXExportChainWithRoot);

            return new X509Certificate2(Convert.FromBase64String(base64encoded), string.Empty, X509KeyStorageFlags.Exportable);
        }

        /// <summary>
        /// Method to find a given certificate.
        /// </summary>
        /// <param name="findType">FindType property</param>
        /// <param name="typeValue">FindType value</param>
        /// <param name="validOnly">Whether to return only valid certificates or not</param>
        /// <returns>Returns a certificate.</returns>
        public static X509Certificate2 GetCertificate(X509FindType findType, string typeValue, bool validOnly)
        {
            X509Certificate2 cert;

            using (X509Store certStore = new X509Store(StoreName.My, StoreLocation.LocalMachine))
            {
                certStore.Open(OpenFlags.ReadOnly);
                X509Certificate2Collection certCollection = certStore.Certificates;
                X509Certificate2Collection result = certCollection.Find(findType, typeValue, validOnly);
                cert = result.Count == 0 ? null : result[0];
            }

            return cert;
        }

        /// <summary>
        /// Gets the TenantId from the config file.
        /// </summary>
        /// <returns>Returns a tenant identifier.</returns>
        public static string GetTenantId()
        {
            string appConfigPath = Uri.UnescapeDataString(new UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path) + ".config";
            XDocument appConfig = XDocument.Load(appConfigPath);
            return appConfig.Descendants("add")
                    .First(node => (string)node.Attribute("key") == "TenantId")
                    .Attribute("value").Value;
        }

        /// <summary>
        /// Gets the ApplicationId from the config file.
        /// </summary>
        /// <returns>Returns an application identifier.</returns>
        public static string GetApplicationId()
        {
            string appConfigPath = Uri.UnescapeDataString(new UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path) + ".config";
            XDocument appConfig = XDocument.Load(appConfigPath);
            return appConfig.Descendants("add")
                    .First(node => (string)node.Attribute("key") == "ApplicationId")
                    .Attribute("value").Value;
        }

        /// <summary>
        /// Gets the CertThumbprint from the config file.
        /// </summary>
        /// <returns>Returns a certificate thumbprint.</returns>
        public static string GetCertThumbprint()
        {
            string appConfigPath = Uri.UnescapeDataString(new UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path) + ".config";
            XDocument appConfig = XDocument.Load(appConfigPath);
            return appConfig.Descendants("add")
                    .First(node => (string)node.Attribute("key") == "CertThumbprint")
                    .Attribute("value").Value;
        }

        /// <summary>
        /// Updates the CertThumbprint in the config file.
        /// </summary>
        /// <param name="newCertThumbprint">Certificate thumbprint</param>
        public static void UpdateCertThumbprint(string newCertThumbprint)
        {
            string appConfigPath = Uri.UnescapeDataString(new UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path) + ".config";
            var appConfig = XElement.Load(appConfigPath);
            var item = appConfig.Descendants("add")
                .First(node => (string)node.Attribute("key") == "CertThumbprint")
                .Attribute("value");
            item.SetValue(newCertThumbprint);
            appConfig.Save(appConfigPath);
        }
    }
}
