using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddInTest
{
	public class ArchiveEmailAttachment
	{
		public String fileName { get; set; }
		public String mediaTypeName { get; set; }
		public byte[] content { get; set; }

		public ArchiveEmailAttachment(String fileName, String mediaTypeName, byte[] content)
		{
			this.fileName = fileName;
			this.mediaTypeName  = mediaTypeName;
			this.content = content;
		}
	}
}
