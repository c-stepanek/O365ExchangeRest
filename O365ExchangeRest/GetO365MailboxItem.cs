//-----------------------------------------------------------------------
// <copyright file="GetO365MailboxItem.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System;
    using System.Management.Automation;
    using Exchange.RestServices;
    using Microsoft.OutlookServices;

    /// <summary>
    /// Defines the <see cref="GetO365MailboxItem"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "O365MailboxItem")]
    public class GetO365MailboxItem : PSCmdlet
    {
        #region Parameters

        /// <summary>
        /// Gets or sets the mailbox SMTP address
        /// </summary>
        [Parameter(ParameterSetName = "FolderName", Mandatory = true)]
        [Parameter(ParameterSetName = "FolderId", Mandatory = true)]
        public string SmtpAddress { get; set; }

        /// <summary>
        /// Gets or sets the mailbox folder name
        /// </summary>
        [Parameter(ParameterSetName = "FolderName", Mandatory = true)]
        public WellKnownFolderName FolderName { get; set; }

        /// <summary>
        /// Gets or sets the mailbox folder identification number
        /// </summary>
        [Parameter(ParameterSetName = "FolderId", Mandatory = true)]
        public FolderId FolderId { get; set; }

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

            MessageView messageView = new MessageView(50);
            FindItemsResults<Item> findItemsResults = null;

            switch (this.ParameterSetName)
            {
                case "FolderName":
                    do
                    {
                        findItemsResults = exchangeService.FindItems(this.FolderName, messageView);
                        messageView.Offset += messageView.PageSize;
                        foreach (Item item in findItemsResults)
                        {
                            this.WriteObject(new PSObject(item));
                        }
                    }
                    while (findItemsResults.MoreAvailable);
                    break;

                case "FolderId":
                    do
                    {
                        findItemsResults = exchangeService.FindItems(this.FolderId, messageView);
                        messageView.Offset += messageView.PageSize;
                        foreach (Item item in findItemsResults)
                        {
                            this.WriteObject(new PSObject(item));
                        }
                    }
                    while (findItemsResults.MoreAvailable);
                    break;

                default:
                    this.ThrowTerminatingError(
                        new ErrorRecord(
                            new ArgumentException("Bad ParameterSetName"), 
                            string.Empty, 
                            ErrorCategory.InvalidOperation, 
                            null));
                return;
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