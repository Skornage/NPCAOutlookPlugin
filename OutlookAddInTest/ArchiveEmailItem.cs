using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddInTest
{
	public class ArchiveEmailItem
	{
		public DateTime WhenReceivedUtc { get; set; }
		public String FromDisplayName { get; set; }
		public String FromEmailAddress { get; set; }
		public String Subject { get; set; }
		public String Body { get; set; }
		public bool IsBodyHtml { get; set; }
		public List<ArchiveEmailAttachment> Attachments { get; set; }

		public ArchiveEmailItem(DateTime whenReceivedUtc, String fromDisplayName, String fromEmailAddress,
			String subject, String body, bool isBodyHtml)
		{
			this.WhenReceivedUtc = whenReceivedUtc;
			this.FromDisplayName = fromDisplayName;
			this.FromEmailAddress = fromEmailAddress;
			this.Subject = subject;
			this.Body = body;
			this.IsBodyHtml = isBodyHtml;
			this.Attachments = new List<ArchiveEmailAttachment>();
		}

		public void addAttachment(ArchiveEmailAttachment attachment)
		{
			this.Attachments.Add(attachment);
		}
	}
}
