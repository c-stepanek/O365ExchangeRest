//-----------------------------------------------------------------------
// <copyright file="GetO365InboxRule.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Management.Automation;
    using Exchange.RestServices;
    using Microsoft.OutlookServices;

    /// <summary>
    /// Defines the <see cref="GetO365InboxRule"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "O365InboxRule")]
    public class GetO365InboxRule : PSCmdlet
    {
        #region Parameters
        
        /// <summary>
        /// Gets or sets the mailbox SMTP address
        /// </summary>
        [Parameter(Mandatory = true)]
        public string SmtpAddress { get; set; }

        /// <summary>
        /// Gets or sets the inbox rule identification number
        /// </summary>
        [Parameter(Mandatory = false)]
        public string RuleId { get; set; }

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
            // Create the Exchange Service Object and auth with bearer token
            ExchangeService exchangeService = new ExchangeService(new AuthenticationProvider(), this.SmtpAddress.ToString(), RestEnvironment.OutlookBeta);

            // If a RuleId was specified search for that, else list all inbox rules
            if (!string.IsNullOrEmpty(this.RuleId))
            {
                this.WriteObject(new PSObject(exchangeService.GetInboxRule(this.RuleId)));
            }
            else
            {
                List<MessageRule> inboxRules = exchangeService.GetInboxRules().ToList();
                inboxRules.ForEach(x => this.WriteObject(new PSObject(x)));
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
