using System;

namespace dotnet_outlook_nosdk.Models
{
  public class OutlookEvent
  {
    public string Organizer { get; set; }
    public string Subject { get; set; }
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    public string LocationDisplayName { get; set; }
    public string LocationAddress { get; set; }
    public string LocationCoordinates { get; set; }
    public string WebLink { get; set; }
    public string OnlineMeetingUrl { get; set; }
  }
}