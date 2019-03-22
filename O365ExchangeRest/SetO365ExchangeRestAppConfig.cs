//-----------------------------------------------------------------------
// <copyright file="SetO365ExchangeRestAppConfig.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Management.Automation;

    /// <summary>
    /// Defines the <see cref="SetO365ExchangeRestAppConfig"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "O365ExchangeRestAppConfig")]
    public class SetO365ExchangeRestAppConfig : PSCmdlet
    {
        #region Parameters

        /// <summary>
        /// Gets or sets the TenantId value.
        /// </summary>
        [Parameter(Mandatory = false)]
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the ApplicationId value.
        /// </summary>
        [Parameter(Mandatory = false)]
        public string ApplicationId { get; set; }

        /// <summary>
        /// Gets or sets the CertThumbprint value.
        /// </summary>
        [Parameter(Mandatory = false)]
        public string CertThumbprint { get; set; }

        #endregion Parameters

        #region Overrides

        /// <summary>
        ///  Performs initialization of command execution
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
        }

        /// <summary>
        /// Provides a record-by-record processing functionality for the cmdlet.
        /// </summary>
        protected override void ProcessRecord()
        {
            PSObject psObject = new PSObject();
            List<KeyValuePair<string, object>> boundParameters = this.MyInvocation.BoundParameters.ToList();

            if (boundParameters.Count >= 1)
            {
                boundParameters.ForEach(
                    param => 
                    {
                        Utils.UpdateConfigKeyValue(param.Key, param.Value.ToString());
                        psObject.Members.Add(new PSNoteProperty(param.Key, Utils.GetConfigKeyValue(param.Key)));
                    });
                this.WriteObject(psObject);
            }
            else
            {
                this.ThrowTerminatingError(
                new ErrorRecord(
                new ArgumentNullException("No parameters were passed."),
                string.Empty,
                ErrorCategory.InvalidOperation,
                null));
            }
        }

        /// <summary>
        /// Performs clean-up after the command execution
        /// </summary>
        protected override void EndProcessing()
        {
            base.EndProcessing();
        }

        #endregion Overrides
    }
}
