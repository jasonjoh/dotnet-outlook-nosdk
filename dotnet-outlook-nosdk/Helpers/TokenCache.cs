using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web;


// This class implements a simplistic "cache" for tokens
// Tokens are mapped by the user's ID (the oid claim from the ID token)
// NOTE: This is storing tokens in a JSON file on disk, which isn't a secure approach.
// This is intended for demonstration purposes only. 

namespace dotnet_outlook_nosdk.Helpers
{
  public class TokenCacheEntry
  {
    public string user_id;
    public string access_token;
    public string refresh_token;
    public string id_token;
    public DateTime expires;
  }

  public static class TokenCache
  {
    // This should work on development machines, but may not work on actual servers due to lack of write permissions
    private static string cacheFileName = HttpContext.Current.Server.MapPath("~/App_Data/token_cache.json");
    private static List<TokenCacheEntry> cacheEntries;

    public static TokenCacheEntry GetUserEntry(string userId)
    {
      LoadCacheFile();

      return cacheEntries.Find(entry => entry.user_id.Equals(userId));
    }

    public static void AddOrUpdateUserEntry(string userId, TokenRequestSuccessResponse data)
    {
      LoadCacheFile();

      TokenCacheEntry userEntry = cacheEntries.Find(entry => entry.user_id.Equals(userId));
      if (userEntry != null)
      {
        userEntry.access_token = data.access_token;
        userEntry.refresh_token = data.refresh_token;
        userEntry.id_token = data.id_token;
        userEntry.expires = DateTime.Now.AddSeconds(Convert.ToDouble(data.expires_in)).AddMinutes(-5);
      }
      else
      {
        cacheEntries.Add(new TokenCacheEntry()
        {
          access_token = data.access_token,
          refresh_token = data.refresh_token,
          id_token = data.id_token,
          expires = DateTime.Now.AddSeconds(Convert.ToDouble(data.expires_in)).AddMinutes(-5),
          user_id = userId
        });
      }

      SaveCacheFile();
    }

    public static void RemoveUserEntry(string userId)
    {
      LoadCacheFile();

      TokenCacheEntry userEntry = cacheEntries.Find(entry => entry.user_id.Equals(userId));

      if (userEntry != null)
      {
        cacheEntries.Remove(userEntry);
      }

      SaveCacheFile();
    }

    private static void LoadCacheFile()
    {
      string cacheContents = string.Empty;
      if (File.Exists(cacheFileName))
      {
        cacheContents = File.ReadAllText(cacheFileName);
      }

      if (!string.IsNullOrEmpty(cacheContents))
      {
        cacheEntries = JsonConvert.DeserializeObject<List<TokenCacheEntry>>(cacheContents);
      }
      else
      {
        cacheEntries = new List<TokenCacheEntry>();
      }
    }

    private static void SaveCacheFile()
    {
      string newCacheContents = JsonConvert.SerializeObject(cacheEntries);
      File.WriteAllText(cacheFileName, newCacheContents);
    }
  }
}