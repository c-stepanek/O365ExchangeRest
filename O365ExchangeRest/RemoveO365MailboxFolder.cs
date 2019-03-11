//-----------------------------------------------------------------------
// <copyright file="RemoveO365MailboxFolder.cs" company="Microsoft Corporation">
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
    /// Defines the <see cref="RemoveO365MailboxFolder"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "O365MailboxFolder")]
    public class RemoveO365MailboxFolder : PSCmdlet
    {
        #region Fields

        /// <summary>
        /// Reference value for ShouldContinue
        /// </summary>
        private bool yesToAll;

        /// <summary>
        /// Reference value for ShouldContinue
        /// </summary>
        private bool noToAll;

        #endregion Fields

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

        /// <summary>
        /// Gets or sets the flag to suppress the ShouldContinue message
        /// </summary>
        [Parameter(Mandatory = false)]
        public SwitchParameter Force { get; set; }

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

            FolderView folderView = new FolderView(10);
            FindFoldersResults findFoldersResults = null;

            do
            {
                SearchFilter searchFilter = new SearchFilter.IsEqualTo(MailFolderObjectSchema.DisplayName, this.FolderName);
                findFoldersResults = exchangeService.FindFolders(this.FolderRoot, searchFilter, folderView);
                folderView.Offset += folderView.PageSize;
                foreach (MailFolder folder in findFoldersResults)
                {
                    if (!this.Force && !this.ShouldContinue(string.Format("This is a hard delete operation. Are you sure you want to delete mail folder {0}?", this.FolderName), "Hard Delete Folder", true, ref this.yesToAll, ref this.noToAll))
                    {
                        return;
                    }

                    // Hard delete the folder
                    folder.Delete();
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
