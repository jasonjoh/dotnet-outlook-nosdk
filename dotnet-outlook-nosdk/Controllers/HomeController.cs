using System;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft.Experimental.IdentityModel.Clients.ActiveDirectory;

using dotnet_outlook_nosdk.Helpers;

namespace dotnet_outlook_nosdk.Controllers
{
  public class HomeController : Controller
  {
    // The Azure login authority
    private static string authority = "https://login.microsoftonline.com/common";
    // The application ID from https://apps.dev.microsoft.com
    private static string appId = System.Configuration.ConfigurationManager.AppSettings["ida:ClientID"];
    // The application secret from https://apps.dev.microsoft.com
    private static string appSecret = System.Configuration.ConfigurationManager.AppSettings["ida:ClientSecret"];

    // The required scopes for our app
    private static string[] scopes = { "https://outlook.office.com/calendars.read" };

    public async Task<ActionResult> Index()
    {

      // By using this version of the AuthenticationContext constructor,
      // we are using the default in-memory token cache. In a real app, you would
      // want to provide an implementation of TokenCache that saves the data somewhere
      // so that you could persist it if restarting the app, etc.
      AuthenticationContext authContext = new AuthenticationContext(authority);

      ClientCredential credential = new ClientCredential(appId, appSecret);
      AuthenticationResult authResult = null;

      ViewBag.Message = TempData["message"];

      try
      {
        authResult = await authContext.AcquireTokenSilentAsync(scopes, credential, UserIdentifier.AnyUser);

        ViewBag.UserName = GetUserEmail(authContext, appId);
      }
      catch (AdalException ex)
      {
        if (ex.ErrorCode == "failed_to_acquire_token_silently")
        {
          // We don't have a token in the cache OR the token couldn't be refreshed
          // We need to have the user sign in
          Uri redirectUri = new Uri(Url.Action("Authorize", "Home", null, Request.Url.Scheme));
          ViewBag.LoginUri = await authContext.GetAuthorizationRequestUrlAsync(scopes, null, appId, redirectUri, UserIdentifier.AnyUser, null);
        }
        else
        {
          TempData["error_message"] = ex.Message;
          RedirectToAction("Error");
        }
      }

      return View();
    }

    public async Task<ActionResult> Authorize()
    {
      string authCode = Request.Params["code"];
      if (string.IsNullOrEmpty(authCode))
      {
        string error = Request.Params["error"];
        string error_description = Request.Params["error_description"];

        TempData["error_message"] = string.Format("Error: {0} - {1}", error, error_description);
        return RedirectToAction("Error");
      }

      AuthenticationContext authContext = new AuthenticationContext(authority);

      ClientCredential credential = new ClientCredential(appId, appSecret);
      AuthenticationResult authResult = null;
      Uri redirectUri = new Uri(Url.Action("Authorize", "Home", null, Request.Url.Scheme));

      try
      {
        authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(authCode, redirectUri, credential, scopes);
      }
      catch (AdalException ex)
      {
        TempData["error_message"] = ex.Message;
        return RedirectToAction("Error");
      }

      return Redirect("/");
    }

    public ActionResult Logout()
    {
      AuthenticationContext authContext = new AuthenticationContext(authority);
      authContext.TokenCache.Clear();
      return Redirect("/");
    }

    public async Task<ActionResult> Inbox()
    {
      // NYI
      return View();
    }

    public async Task<ActionResult> Calendar()
    {
      AuthenticationContext authContext = new AuthenticationContext(authority);

      ClientCredential credential = new ClientCredential(appId, appSecret);
      AuthenticationResult authResult = null;
      try
      {
        authResult = await authContext.AcquireTokenSilentAsync(scopes, credential, UserIdentifier.AnyUser);
      }
      catch (AdalException ex)
      {
        TempData["message"] = "Please sign in to continue";
        return Redirect("/");
      }

      var client = new Outlook();
      client.anchorMailbox = GetUserEmail(authContext, appId);
      ViewBag.UserName = client.anchorMailbox;

      DateTime viewStart = DateTime.Now.Date.ToUniversalTime();
      DateTime viewEnd = viewStart.AddDays(7);
      var result = await client.GetCalendarView(authResult.Token, client.anchorMailbox, viewStart, viewEnd);

      return View(result);
    }

    public async Task<ActionResult> Contacts()
    {
      // NYI
      return View();
    }

    public ActionResult Error()
    {
      ViewBag.ErrorMessage = TempData["error_message"];
      return View();
    }

    private string GetUserEmail(AuthenticationContext context, string clientId)
    {
      // ADAL caches the ID token in its token cache by the client ID
      foreach (TokenCacheItem item in context.TokenCache.ReadItems())
      {
        if (item.Scope.Contains(clientId))
        {
          return GetEmailFromIdToken(item.Token);
        }
      }
      return string.Empty;
    }

    private string GetEmailFromIdToken(string token)
    {
      // JWT is made of three parts, separated by a '.' 
      // First part is the header 
      // Second part is the token 
      // Third part is the signature 
      string[] tokenParts = token.Split('.');
      if (tokenParts.Length < 3)
      {
        // Invalid token, return empty
      }
      // Token content is in the second part, in urlsafe base64
      string encodedToken = tokenParts[1];
      // Convert from urlsafe and add padding if needed
      int leftovers = encodedToken.Length % 4;
      if (leftovers == 2)
      {
        encodedToken += "==";
      }
      else if (leftovers == 3)
      {
        encodedToken += "=";
      }
      encodedToken = encodedToken.Replace('-', '+').Replace('_', '/');
      // Decode the string
      var base64EncodedBytes = System.Convert.FromBase64String(encodedToken);
      string decodedToken = System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
      // Load the decoded JSON into a dynamic object
      dynamic jwt = Newtonsoft.Json.JsonConvert.DeserializeObject(decodedToken);
      // User's email is in the preferred_username field
      return jwt.preferred_username;
    }
  }
}