//-----------------------------------------------------------------------
// <copyright file="NewO365MailboxMessage.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System.Collections.Generic;
    using System.Management.Automation;
    using Exchange.RestServices;
    using Microsoft.OutlookServices;

    /// <summary>
    /// Defines the <see cref="NewO365MailboxMessage"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.New, "O365MailboxMessage")]
    public class NewO365MailboxMessage : PSCmdlet
    {
        #region Parameters

        /// <summary>
        /// Gets or sets the mailbox SMTP address
        /// </summary>
        [Parameter(Mandatory = true)]
        public string SmtpAddress { get; set; }

        /// <summary>
        /// Gets or sets the message subject
        /// </summary>
        [Parameter(Mandatory = true)]
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets the message body
        /// </summary>
        [Parameter(Mandatory = true)]
        public string Body { get; set; }

        /// <summary>
        /// Gets or sets the To recipients
        /// </summary>
        [Parameter(Mandatory = true)]
        public List<string> Recipients { get; set; }

        /// <summary>
        /// Gets or sets the flag to send the message
        /// </summary>
        [Parameter(Mandatory = false)]
        public SwitchParameter SendMessage { get; set; }

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

            List<Recipient> recipientList = new List<Recipient>();
            this.Recipients.ForEach(recipient => recipientList.Add(new Recipient { EmailAddress = new EmailAddress { Address = recipient } }));

            Message message = new Message(exchangeService)
            {
                Subject = this.Subject,
                Body = new ItemBody()
                {
                    ContentType = BodyType.HTML,
                    Content = this.Body
                },
                ToRecipients = recipientList,
                From = new Recipient()
                {
                    EmailAddress = new EmailAddress()
                    {
                        Address = this.SmtpAddress
                    }
                }
            };

            if (this.SendMessage)
            {
                message.Send();
            }
            else
            {
                message.Save(WellKnownFolderName.Drafts);
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