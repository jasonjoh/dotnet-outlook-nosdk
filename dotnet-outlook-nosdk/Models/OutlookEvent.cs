using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace dotnet_outlook_nosdk.Models
{
  public class OutlookEvent
  {
    public string Organizer { get; set; }
    public string Subject { get; set; }
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    public string Location { get; set; }
    public string WebLink { get; set; }
  }
}