//-----------------------------------------------------------------------
// <copyright file="GetO365MailboxFolder.cs" company="Microsoft Corporation">
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
    /// Defines the <see cref="GetO365MailboxFolder"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "O365MailboxFolder")]
    public class GetO365MailboxFolder : PSCmdlet
    {
        #region Parameters

        /// <summary>
        /// Gets or sets the mailbox SMTP address
        /// </summary>
        [Parameter(Mandatory = true)]
        public string SmtpAddress { get; set; }

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

            FolderView folderView = new FolderView(10);
            FindFoldersResults findFoldersResults = null;

            do
            {
                findFoldersResults = exchangeService.FindFolders(WellKnownFolderName.MsgFolderRoot, folderView);
                folderView.Offset += folderView.PageSize;
                foreach (MailFolder folder in findFoldersResults)
                {
                    this.WriteObject(new PSObject(folder));
                }
            }
            while (findFoldersResults.MoreAvailable);
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
