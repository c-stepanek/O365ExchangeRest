//-----------------------------------------------------------------------
// <copyright file="GetO365MailboxFolder.cs" company="Microsoft Corporation">
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
    /// Defines the <see cref="GetO365MailboxFolder"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "O365MailboxFolder")]
    public class GetO365MailboxFolder : PSCmdlet
    {
        #region Parameters

        /// <summary>
        /// Gets or sets the mailbox SMTP address
        /// </summary>
        [Parameter(ParameterSetName = "Default", Mandatory = true)]
        [Parameter(ParameterSetName = "FolderName", Mandatory = true)]
        [Parameter(ParameterSetName = "FolderId", Mandatory = true)]
        public string SmtpAddress { get; set; }

        /// <summary>
        /// Gets or sets the mailbox folder name
        /// </summary>
        [Parameter(ParameterSetName = "FolderName", Mandatory = true)]
        public string FolderName { get; set; }

        /// <summary>
        /// Gets or sets the mailbox folder identification number
        /// </summary>
        [Parameter(ParameterSetName = "FolderId", Mandatory = true)]
        public FolderId FolderId { get; set; }

        /// <summary>
        /// Gets or sets the mailbox folder root
        /// </summary>
        [Parameter(ParameterSetName = "Default", Mandatory = false)]
        [Parameter(ParameterSetName = "FolderName", Mandatory = false)]
        [Parameter(ParameterSetName = "FolderId", Mandatory = false)]
        public WellKnownFolderName FolderRoot { get; set; } = WellKnownFolderName.MsgFolderRoot;

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

            switch (this.ParameterSetName)
            {
                case "FolderName":
                    do
                    {
                        SearchFilter searchFilter = new SearchFilter.IsEqualTo(MailFolderObjectSchema.DisplayName, this.FolderName);
                        findFoldersResults = exchangeService.FindFolders(this.FolderRoot, searchFilter, folderView);
                        folderView.Offset += folderView.PageSize;
                        foreach (MailFolder folder in findFoldersResults)
                        {
                            this.WriteObject(new PSObject(folder));
                        }
                    }
                    while (findFoldersResults.MoreAvailable);
                    break;

                case "FolderId":
                    do
                    {
                        findFoldersResults = exchangeService.FindFolders(this.FolderRoot, folderView);
                        folderView.Offset += folderView.PageSize;
                        foreach (MailFolder folder in findFoldersResults)
                        {
                            if (folder.Id == this.FolderId.ToString())
                            {
                                this.WriteObject(new PSObject(folder));
                            }
                        }
                    }
                    while (findFoldersResults.MoreAvailable);
                    break;

                default:
                    do
                    {
                        findFoldersResults = exchangeService.FindFolders(this.FolderRoot, folderView);
                        folderView.Offset += folderView.PageSize;
                        foreach (MailFolder folder in findFoldersResults)
                        {
                            this.WriteObject(new PSObject(folder));
                        }
                    }
                    while (findFoldersResults.MoreAvailable);
                    break;
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
