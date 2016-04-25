using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace dotnet_outlook_nosdk.Helpers
{
  public class OAuthHelper
  {
    private static string authEndpoint = "/oauth2/v2.0/authorize";
    private static string tokenEndpoint = "/oauth2/v2.0/token";

    public string Authority { get; set; }
    public string AppId { get; set; }
    public string AppSecret { get; set; }

    public OAuthHelper(string authority, string appId, string appSecret)
    {
      Authority = authority;
      AppId = appId;
      AppSecret = appSecret;
    }

    // Builds the authorization URL where the app sends the user to sign in
    public string GetAuthorizationUrl(string[] scopes, string redirectUri, string state, string nonce)
    {
      // Start with the base URL
      UriBuilder authUrl = new UriBuilder(this.Authority + authEndpoint);

      authUrl.Query =
        "response_type=code+id_token" +
        "&scope=openid+offline_access+profile+" + GetEncodedScopes(scopes) +
        "&state=" + state +
        "&nonce=" + nonce + 
        "&client_id=" + this.AppId +
        "&redirect_uri=" + HttpUtility.UrlEncode(redirectUri) +
        "&response_mode=form_post";

      return authUrl.ToString();
    }

    // Makes a POST request to the token endopoint to get an access token using either
    // an authorization code or a refresh token. This will also add the tokens
    // to the local cache.
    public async Task<TokenRequestSuccessResponse> GetTokensFromAuthority(string grantType, string grantParameter, string redirectUri, string userId)
    {
      // Build the token request payload
      FormUrlEncodedContent tokenRequestForm;

      if (grantType.Equals("authorization_code"))
      {
        tokenRequestForm = new FormUrlEncodedContent(
          new[] {
            new KeyValuePair<string,string>("grant_type", "authorization_code"),
            new KeyValuePair<string,string>("code", grantParameter),
            new KeyValuePair<string,string>("client_id", this.AppId),
            new KeyValuePair<string,string>("client_secret", this.AppSecret),
            new KeyValuePair<string,string>("redirect_uri", redirectUri)
          }
        );
      }
      else
      {
        tokenRequestForm = new FormUrlEncodedContent(
          new[] {
            new KeyValuePair<string,string>("grant_type", "refresh_token"),
            new KeyValuePair<string,string>("code", grantParameter),
            new KeyValuePair<string,string>("client_id", this.AppId),
            new KeyValuePair<string,string>("client_secret", this.AppSecret),
            new KeyValuePair<string,string>("redirect_uri", redirectUri)
          }
        );
      }

      using (HttpClient httpClient = new HttpClient())
      {
        string requestString = tokenRequestForm.ReadAsStringAsync().Result;
        StringContent requestContent = new StringContent(requestString);
        requestContent.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

        HttpRequestMessage tokenRequest = new HttpRequestMessage(HttpMethod.Post, this.Authority + tokenEndpoint);
        tokenRequest.Content = requestContent;
        tokenRequest.Headers.UserAgent.Add(new ProductInfoHeaderValue("dotnet-outlook-nosdk", "1.0"));
        tokenRequest.Headers.Add("client-request-id", Guid.NewGuid().ToString());
        tokenRequest.Headers.Add("return-client-request-id", "true");

        HttpResponseMessage response = await httpClient.SendAsync(tokenRequest);
        JObject jsonResponse = JObject.Parse(response.Content.ReadAsStringAsync().Result);
        JsonSerializer jsonSerializer = new JsonSerializer();

        if (response.IsSuccessStatusCode)
        {
          TokenRequestSuccessResponse s = (TokenRequestSuccessResponse)jsonSerializer.Deserialize(new JTokenReader(jsonResponse), typeof(TokenRequestSuccessResponse));
          TokenCache.AddOrUpdateUserEntry(userId, s);
          return s;
        }
        else
        {
          TokenRequestErrorResponse e = (TokenRequestErrorResponse)jsonSerializer.Deserialize(new JTokenReader(jsonResponse), typeof(TokenRequestErrorResponse));
          throw new Exception(e.error_description);
        }
      }
    }

    // Attempts to read the user's token from the cache. If the token is expired or close to expiring,
    // it will refresh the token, update the cache, and return the new token. Returns null if not found
    // or cannot refresh.
    public async Task<string> GetAccessToken(string userId, string redirectUri)
    {
      TokenCacheEntry userEntry = TokenCache.GetUserEntry(userId);

      if (userEntry == null)
      {
        return null;
      }

      if (userEntry.expires.CompareTo(DateTime.Now.AddMinutes(5)) < 0)
      {
        if (string.IsNullOrEmpty(userEntry.refresh_token))
        {
          return null;
        }

        // Refresh the token
        TokenRequestSuccessResponse response = await GetTokensFromAuthority("refresh_token", userEntry.refresh_token, redirectUri, userId);
        return response.access_token;
      }
      else
      {
        return userEntry.access_token;
      }
    }

    // Removes the user's token from the cache
    public void LogOut(string userId)
    {
      TokenCache.RemoveUserEntry(userId);
    }

    private string GetEncodedScopes(string[] scopes)
    {
      string encodedScopes = string.Empty;
      foreach(string scope in scopes)
      {
        if (!string.IsNullOrEmpty(encodedScopes)) { encodedScopes += '+'; }
        encodedScopes += HttpUtility.UrlEncode(scope);
      }
      return encodedScopes;
    }
  }
}