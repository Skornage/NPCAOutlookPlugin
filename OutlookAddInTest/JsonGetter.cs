using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;

namespace OutlookAddInTest
{
	class JsonGetter
	{
<<<<<<< HEAD
		public static JObject getJsonObject()
		{
			//getting companyID, id, type, name, email
			return JObject.Parse(@"
			{
				""jagged"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged1"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""Josh"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagge2"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""Noah"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged3"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""Alex"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged4"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""New Name"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged5"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""Old Name"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagge6"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""Searching"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged7"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""Other"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged8"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""Jason"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged9"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged10"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged11"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged12"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged13"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged14"":
					{
						""id"":""value2"",
						""type"":""Contact"",
						""name"":""Jason"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged15"":
					{
						""id"":""value2"",
						""type"":""Contact"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged16"":
					{
						""id"":""value2"",
						""type"":""Contact"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged17"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged18"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged19"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged20"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged21"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged22"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged2"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged2"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged2"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged2"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged2"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged2"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged2"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					},
				""jagged3"":
					{
						""id"":""value2"",
						""type"":""Account"",
						""name"":""NPCA"",
						""email"":""millerna@rose-hulman.edu"",
						""info"":""test""
					}
			}");
		}

=======
		public static List<Result> GetData()
        {
			List<Result> model = null;
			var client = new HttpClient();
			var task = client.GetAsync("http://phoenix-dev.azurewebsites.net/api/v1/outlook/businessentities?apiToken=MUg@R*A8jgtwY$aQXv3J")
			  .ContinueWith((taskwithresponse) =>
			  {
				  var response = taskwithresponse.Result;
				  var jsonString = response.Content.ReadAsStringAsync();
				  jsonString.Wait();
				  model = JsonConvert.DeserializeObject<List<Result>>(jsonString.Result);
			  });
			task.Wait();
			return model;
        }
>>>>>>> 42092be9fabdb0688e725028ce364388fd39d67f
	}

	public class Result
	{
		public string id { get; set; }
		public int idNumber { get; set; }
		public bool isInactive { get; set; }
		public bool isContact { get; set; }
		public bool isCompany { get; set; }
		public string companyType { get; set; }
		public bool hasParentCompany { get; set; }
		public bool isPendingMember { get; set; }
		public bool isCurrentMember { get; set; }
		public bool isExpiredMember { get; set; }
		public string membershipType { get; set; }
		public string name { get; set; }
		public object phoneNumber { get; set; }
		public object extension { get; set; }
		public string street { get; set; }
		public string city { get; set; }
		public string stateProvince { get; set; }
		public string postalCode { get; set; }
		public string country { get; set; }
		public string websiteUrl { get; set; }
		public string emailAddress { get; set; }
	}
}
