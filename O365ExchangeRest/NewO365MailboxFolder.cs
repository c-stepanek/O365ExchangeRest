//-----------------------------------------------------------------------
// <copyright file="NewO365MailboxFolder.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System.Management.Automation;
    using Exchange.RestServices;
    using Microsoft.OutlookServices;

    /// <summary>
    /// Defines the <see cref="NewO365MailboxFolder"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.New, "O365MailboxFolder")]
    public class NewO365MailboxFolder : PSCmdlet
    {
        #region Parameters

        /// <summary>
        /// Gets or sets the Param1 value
        /// </summary>
        [Parameter(Mandatory = true)]
        public string SmtpAddress { get; set; }

        /// <summary>
        /// Gets or sets the folder name
        /// </summary>
        [Parameter(Mandatory = true)]
        public string FolderName { get; set; }

        /// <summary>
        /// Gets or sets the mailbox folder root
        /// </summary>
        [Parameter(Mandatory = true)]
        public WellKnownFolderName FolderRoot { get; set; }

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
            ExchangeService exchangeService = new ExchangeService(new AuthenticationProvider(), this.SmtpAddress, RestEnvironment.OutlookBeta);

            MailFolder mailFolder = new MailFolder(exchangeService)
            {
                DisplayName = this.FolderName
            };
            mailFolder.Save(this.FolderRoot);

            FolderView folderView = new FolderView(10);
            FindFoldersResults findFoldersResults = null;

            do
            {
                // Find the new folder and output the result
                SearchFilter searchFilter = new SearchFilter.IsEqualTo(MailFolderObjectSchema.DisplayName, mailFolder.DisplayName);
                findFoldersResults = exchangeService.FindFolders(this.FolderRoot, searchFilter, folderView);
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
