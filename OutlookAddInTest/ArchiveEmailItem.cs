using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddInTest
{
	public class ArchiveEmailItem
	{
		public DateTime whenReceivedUtc { get; set; }
		public String fromDisplayName { get; set; }
		public String fromEmailAddress { get; set; }
		public String subject { get; set; }
		public String body { get; set; }
		public bool isBodyHtml { get; set; }
		public List<ArchiveEmailAttachment> attachments { get; set; }

		public ArchiveEmailItem(DateTime whenReceivedUtc, String fromDisplayName, String fromEmailAddress,
			String subject, String body, bool isBodyHtml)
		{
			this.whenReceivedUtc = whenReceivedUtc;
			this.fromDisplayName = fromDisplayName;
			this.fromEmailAddress = fromEmailAddress;
			this.subject = subject;
			this.body = body;
			this.isBodyHtml = isBodyHtml;
			this.attachments = new List<ArchiveEmailAttachment>();
		}

		public void addAttachment(ArchiveEmailAttachment attachment)
		{
			this.attachments.Add(attachment);
		}
	}
}
