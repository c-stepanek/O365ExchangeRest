//-----------------------------------------------------------------------
// <copyright file="NewO365ExchangeServiceCertificate.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System;
    using System.Management.Automation;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Principal;
    using System.Threading;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Defines the <see cref="NewO365ExchangeServiceCertificate"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.New, "O365ExchangeServiceCertificate")]
    public class NewO365ExchangeServiceCertificate : PSCmdlet
    {
        #region Parameters

        /// <summary>
        /// Gets or sets the SubjectName for the self-signed certificate
        /// </summary>
        [Parameter(Mandatory = true)]
        public string Name { get; set; }

        #endregion Parameters

        #region Overrides

        /// <summary>
        ///  Performs initialization of command execution
        /// </summary>
        protected override void BeginProcessing()
        {
            AppDomain currentDomain = Thread.GetDomain();

            currentDomain.SetPrincipalPolicy(PrincipalPolicy.WindowsPrincipal);
            WindowsPrincipal currentPrincipal = (WindowsPrincipal)Thread.CurrentPrincipal;
            bool runAsAdmin = currentPrincipal.IsInRole(WindowsBuiltInRole.Administrator);

            if (!runAsAdmin)
            {
                this.ThrowTerminatingError(
                    new ErrorRecord(
                        new Exception("Please run this command in an elevated prompt."), 
                        string.Empty, 
                        ErrorCategory.PermissionDenied, 
                        null));
            }
        }

        /// <summary>
        /// Provides a record-by-record processing functionality for the cmdlet.
        /// </summary>
        protected override void ProcessRecord()
        {
            X509Certificate2 cert = Utils.GetCertificate(X509FindType.FindBySubjectName, this.Name, false);

            if (cert == null)
            {
                // Create new self-signed certificate
                cert = Utils.NewSelfSignedCertificate(this.Name);
            }

            Manifest manifest = new Manifest
            {
                keyId = Guid.NewGuid(),
                value = Convert.ToBase64String(cert.RawData),
                type = "AsymmetricX509Cert",
                usage = "Verify",
                customKeyIdentifier = Convert.ToBase64String(cert.GetCertHash())
            };

            this.WriteHost("\nIf neccessary, certificate can be exported as pfx by running following command:");
            this.WriteHost("\t\t$certificatePassword = Read-Host -Prompt 'Enter pfx password' -AsSecureString");
            this.WriteHost($"\t\tExport-PfxCertificate -Cert 'Cert:\\LocalMachine\\My\\{cert.Thumbprint}' -FilePath 'C:\\Temp\\cert.pfx' -Password $certificatePassword\n");

            string json = JsonConvert.SerializeObject(manifest);
            string formattedJson = JObject.Parse(json).ToString();
            this.WriteObject(formattedJson);
        }

        /// <summary>
        /// Performs clean-up after the command execution
        /// </summary>
        protected override void EndProcessing()
        {
            base.EndProcessing();
        }

        #endregion Overrides

        #region Private Methods

        /// <summary>
        /// Simulate Write-Host behavior
        /// </summary>
        /// <param name="message">Test to display in host UI</param>
        private void WriteHost(string message)
        {
            if (this.CommandRuntime.Host != null)
            {
                this.CommandRuntime.Host.UI.WriteLine(message);
            }
        }

        #endregion Private Methods
    }
}
