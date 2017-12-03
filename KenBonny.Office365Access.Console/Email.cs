using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace KenBonny.Office365Access.Console
{
    internal class EmailCollection
    {
        [JsonProperty("value")]
        public IReadOnlyCollection<Email> Emails { get; set; }
    }
    
    internal class Email
    {

        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }

        [JsonProperty("@odata.id")]
        public string OdataId { get; set; }

        [JsonProperty("@odata.etag")]
        public string OdataEtag { get; set; }

        [JsonProperty("Id")]
        public string Id { get; set; }

        [JsonProperty("CreatedDateTime")]
        public DateTime CreatedDateTime { get; set; }

        [JsonProperty("LastModifiedDateTime")]
        public DateTime LastModifiedDateTime { get; set; }

        [JsonProperty("ChangeKey")]
        public string ChangeKey { get; set; }

        [JsonProperty("Categories")]
        public string[] Categories { get; set; }

        [JsonProperty("ReceivedDateTime")]
        public DateTime ReceivedDateTime { get; set; }

        [JsonProperty("SentDateTime")]
        public DateTime SentDateTime { get; set; }

        [JsonProperty("HasAttachments")]
        public bool HasAttachments { get; set; }

        [JsonProperty("Subject")]
        public string Subject { get; set; }

        [JsonProperty("Body")]
        public Body Body { get; set; }

        [JsonProperty("BodyPreview")]
        public string BodyPreview { get; set; }

        [JsonProperty("Importance")]
        public string Importance { get; set; }

        [JsonProperty("ParentFolderId")]
        public string ParentFolderId { get; set; }

        [JsonProperty("Sender")]
        public Contact Sender { get; set; }

        [JsonProperty("From")]
        public Contact From { get; set; }

        [JsonProperty("ToRecipients")]
        public Contact[] ToRecipients { get; set; }

        [JsonProperty("CcRecipients")]
        public Contact[] CcRecipients { get; set; }

        [JsonProperty("BccRecipients")]
        public Contact[] BccRecipients { get; set; }

        [JsonProperty("ReplyTo")]
        public Contact[] ReplyTo { get; set; }

        [JsonProperty("ConversationId")]
        public string ConversationId { get; set; }

        [JsonProperty("IsDeliveryReceiptRequested")]
        public bool IsDeliveryReceiptRequested { get; set; }

        [JsonProperty("IsReadReceiptRequested")]
        public bool IsReadReceiptRequested { get; set; }

        [JsonProperty("IsRead")]
        public bool IsRead { get; set; }

        [JsonProperty("IsDraft")]
        public bool IsDraft { get; set; }

        [JsonProperty("WebLink")]
        public string WebLink { get; set; }
    }
    
    internal class Contact
    {
        [JsonProperty("EmailAddress")]
        public EmailAddress EmailAddress { get; set; }
    }
    
    internal class EmailAddress
    {

        [JsonProperty("Name")]
        public string Name { get; set; }

        [JsonProperty("Address")]
        public string Address { get; set; }
    }
    
    internal class Body
    {

        [JsonProperty("ContentType")]
        public string ContentType { get; set; }

        [JsonProperty("Content")]
        public string Content { get; set; }
    }
}