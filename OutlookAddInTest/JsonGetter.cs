using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;

namespace OutlookAddInTest
{
	class JsonGetter
	{
		public static List<Result> GetData()
        {
			List<Result> model = null;
			var client = new HttpClient();
			var task = client.GetAsync("https://npca-phoenix-staging.azurewebsites.net/api/v1/outlook/businessentities?apiToken=MUg@R*A8jgtwY$aQXv3J")
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

		public static List<Result> GetSearchResults(String[] query)
		{
			List<Result> model = null;
			var client = new HttpClient();
			JsonSearch jsonSearch = new JsonSearch();
			jsonSearch.searchTerms = query;
			var jsonSerializerSettings = new JsonSerializerSettings
			{
				ContractResolver = new CamelCasePropertyNamesContractResolver()
			};
			var searchTerms = new StringContent(JsonConvert.SerializeObject(jsonSearch, jsonSerializerSettings), Encoding.UTF8, "application/json");

			try
			{
				var task = client.PostAsync("https://portal.precast.org/api/v1/outlook/businessentities/search?apiToken=MUg@R*A8jgtwY$aQXv3J", searchTerms)
					  .ContinueWith((taskwithresponse) =>
					  {
						  var response = taskwithresponse.Result;
						  var jsonString = response.Content.ReadAsStringAsync();
						  jsonString.Wait();
						  model = JsonConvert.DeserializeObject<List<Result>>(jsonString.Result);
					  });
				task.Wait();
			}
			catch
			{
				System.Windows.Forms.MessageBox.Show("Failed to retrive information from the server. Please try again.");
			}
			return model;
		}
	}

	public class JsonSearch
	{
		public string[] searchTerms;
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
