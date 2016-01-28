using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json.Linq;

using dotnet_outlook_nosdk.Models;

namespace dotnet_outlook_nosdk.Helpers
{
  public class Outlook
  {
    // Used to set the base API endpoint, e.g. "https://outlook.office.com/api/beta"
    public string apiEndpoint { get; set; }
    public string anchorMailbox { get; set; }

    public Outlook()
    {
      // Set default endpoint
      apiEndpoint = "https://outlook.office.com/api/beta";
      anchorMailbox = string.Empty;
    }

    public async Task<HttpResponseMessage> MakeApiCall(string method, string token, string apiUrl, string userEmail, string payload, Dictionary<string, string> preferHeaders)
    {
      using (var httpClient = new HttpClient())
      {
        var request = new HttpRequestMessage(new HttpMethod(method), apiUrl);

        // Headers
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        request.Headers.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("dotnet-outlook-nosdk", "1.0"));
        request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
        request.Headers.Add("client-request-id", Guid.NewGuid().ToString());
        request.Headers.Add("return-client-request-id", "true");
        request.Headers.Add("X-AnchorMailbox", userEmail);
        
        if (preferHeaders != null)
        {
          foreach(KeyValuePair<string, string> header in preferHeaders)
          {
            request.Headers.Add("Prefer", string.Format("{0}=\"{1}\"", header.Key, header.Value));
          }
        }

        // Content
        if ((method.ToUpper() == "POST" || method.ToUpper() == "PATCH") &&
            !string.IsNullOrEmpty(payload))
        {
          request.Content = new StringContent(payload);
          request.Content.Headers.ContentType.MediaType = "application/json";
        }

        var apiResult = await httpClient.SendAsync(request);
        return apiResult;
      }
    }

    public async Task<object> GetCalendarView(string token, string userEmail, DateTime viewStart, DateTime viewEnd)
    {
      string getCalendarViewEndpoint = this.apiEndpoint + "/Me/CalendarView";
      string query = "?startdatetime={0}&enddatetime={1}&$orderby=Start/DateTime&$select=Subject,Organizer,Start,End,Location,WebLink,OnlineMeetingUrl";
      getCalendarViewEndpoint += string.Format(query, viewStart.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"), viewEnd.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"));

      Dictionary<string, string> preferences = new Dictionary<string, string>();
      preferences.Add("exchange.behavior", "onlinemeeting");

      var result = await MakeApiCall("GET", token, getCalendarViewEndpoint, userEmail, null, preferences);

      var response = await result.Content.ReadAsStringAsync();

      // The response looks like:
      // {
      //   @odata.context = "...",
      //   @odata.nextLink = "...",
      //   value: [
      //     {
      //       <Event 1>
      //     },
      //     {
      //       <Event 2>
      //     },
      //     ...
      //   ]

      // Use Json.Net's LINQ functions to get to the "value", which is an array
      // Then we can use LINQ to deserialize to the OutlookEvent class

      JObject responseJson = JObject.Parse(response);
      JArray eventJson = (JArray)responseJson["value"];

      List<OutlookEvent> events = eventJson.Select(e => new OutlookEvent
      {
        Subject = (string)e["Subject"],
        Organizer = BuildOrganizerString(e["Organizer"]["EmailAddress"]),
        Start = DateTime.Parse((string)e["Start"]["DateTime"]),
        End = DateTime.Parse((string)e["End"]["DateTime"]),
        LocationDisplayName = (string)e["Location"]["DisplayName"],
        LocationAddress = BuildAddressString(e["Location"]["Address"]),
        LocationCoordinates = BuildCoordinatesString(e["Location"]["Coordinates"]),
        WebLink = (string)e["WebLink"],
        OnlineMeetingUrl = (string)e["OnlineMeetingUrl"]
      }).ToList();
        
      return events;
    }
    
    public String BuildOrganizerString(JToken emailAddress)
    {
        return String.Format("{0} <{1}>",
            (string)emailAddress["Name"] ?? "<No Name>",
            (string)emailAddress["Address"]
            );
    }

    public String BuildAddressString(JToken address)
    {
        if (address == null)
        {
            return "null";
        }
        else {
            return String.Format("{0}, {1}, {2}, {3}, {4}",
                (string)address["Street"] == "" ? "<No Street>" : (string)address["Street"],
                (string)address["City"] == "" ? "<No City>" : (string)address["City"],
                (string)address["State"] == "" ? "<No State>" : (string)address["State"],
                (string)address["CountryOrRegion"] == "" ? "<No Country Or Region>" : (string)address["CountryOrRegion"],
                (string)address["PostalCode"] == "" ? "<No Postal Code>" : (string)address["PostalCode"]
                );
        }
    }

    public String BuildCoordinatesString(JToken coordinates)
    {
        if (coordinates == null)
        {
            return "null";
        }
        else
        {
            return String.Format("{0}, {1}", (string)coordinates["Latitude"], (string)coordinates["Longitude"]);
        }
    }
  }
}