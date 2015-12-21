using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddInTest
{
	public class ArchiveEmailAttachment
	{
		public String FileName { get; set; }
		public String MediaTypeName { get; set; }
		public byte[] Content { get; set; }

		public ArchiveEmailAttachment(String fileName, String mediaTypeName, byte[] content)
		{
			this.FileName = fileName;
			this.MediaTypeName  = mediaTypeName;
			this.Content = content;
		}
	}
}
