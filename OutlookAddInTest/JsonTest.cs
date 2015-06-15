using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddInTest
{
	class JsonTest
	{
		public static String getJsonObject()
		{
			return Newtonsoft.Json.JsonConvert.SerializeObject(new
			{
				results = new List<Result>()
				{
					new Result { id = 1, type = "Account", name = "NPCA", email = "ABC", info = "ABC" },
					new Result { id = 2, type = "Contact", name = "Blah", email = "JKL", info = "JKL" }
				}
			});
		}

	}

	public class Result
	{
		public int id { get; set; }
		public String type { get; set; }
		public String name {get; set; }
		public string email { get; set; }
		public string info { get; set; }
	}
}
