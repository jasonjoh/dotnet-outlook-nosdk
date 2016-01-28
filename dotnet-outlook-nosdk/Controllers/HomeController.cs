using System;
using System.Threading.Tasks;
using System.Web.Mvc;

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
      // If any message was returned, add it to the ViewBag
      ViewBag.Message = TempData["message"];

      string redirectUri = Url.Action("Authorize", "Home", null, Request.Url.Scheme);
      OAuthHelper oauthHelper = new OAuthHelper(authority, appId, appSecret);
      if (Session["user_name"] != null && Session["user_id"] != null)
      {
        // Make sure token is still good
        try
        {
          string token = await oauthHelper.GetAccessToken((string)Session["user_id"], redirectUri);

          if (!string.IsNullOrEmpty(token))
          {
            ViewBag.UserName = (string)Session["user_name"];
            return View();
          }
        }
        catch (Exception)
        {
          // Clear session and have user login again
          Session.Remove("user_name");
          Session.Remove("user_id");
        }
      }

      // We don't have a token in the cache OR the token couldn't be refreshed
      // We need to have the user sign in
        
      string state = Guid.NewGuid().ToString();
      string nonce = Guid.NewGuid().ToString();
      Session["auth_state"] = state;
      Session["auth_nonce"] = nonce;

      ViewBag.LoginUri = oauthHelper.GetAuthorizationUrl(scopes, redirectUri, state, nonce);

      return View();
    }

    public async Task<ActionResult> Authorize()
    {
      string authState = Request.Params["state"];
      string expectedState = (string)Session["auth_state"];
      Session.Remove("auth_state");

      // Make sure that the state passed by the caller matches what we expect
      if (!authState.Equals(expectedState))
      {
        TempData["error_message"] = "The auth state did not match the expected value. Please try again.";
        return RedirectToAction("Error");
      }

      string authCode = Request.Params["code"];
      string idToken = Request.Params["id_token"];

      // Make sure we got back an auth code and ID token
      if (string.IsNullOrEmpty(authCode) || string.IsNullOrEmpty(idToken))
      {
        string error = Request.Params["error"];
        string error_description = Request.Params["error_description"];

        if (string.IsNullOrEmpty(error) && string.IsNullOrEmpty(error_description))
        {
          TempData["error_message"] = "Missing authorization code and/or ID token from redirect.";
        }
        else
        {
          TempData["error_message"] = string.Format("Error: {0} - {1}", error, error_description);
        }

        return RedirectToAction("Error");
      }

      // Check the nonce in the ID token against what we expect
      string nonce = (string)Session["auth_nonce"];
      Session.Remove("auth_nonce");
      if (!OpenIdToken.ValidateOpenIdToken(idToken, nonce))
      {
        TempData["error_message"] = "Invalid ID token.";
        return RedirectToAction("Error");
      }

      OpenIdToken userId = OpenIdToken.ParseOpenIdToken(idToken);

      OAuthHelper oauthHelper = new OAuthHelper(authority, appId, appSecret);
      string redirectUri = Url.Action("Authorize", "Home", null, Request.Url.Scheme);
      try
      {
        TokenRequestSuccessResponse response = await oauthHelper.GetTokensFromAuthority("authorization_code", authCode, redirectUri, userId.oid);

        Session["user_name"] = GetEmailFromIdToken(response.id_token);
        Session["user_id"] = userId.oid;
      }
      catch (Exception ex)
      {
        TempData["error_message"] = ex.Message;
        return RedirectToAction("Error");
      }

      return Redirect("/");
    }

    public ActionResult Logout()
    {
      OAuthHelper oauthHelper = new OAuthHelper(authority, appId, appSecret);
      oauthHelper.LogOut((string)Session["user_id"]);

      Session.Remove("user_name");
      Session.Remove("user_id");

      return Redirect("/");
    }

    public async Task<ActionResult> Inbox()
    {
      // NYI
      return View();
    }

    public async Task<ActionResult> Calendar()
    {
      string userId = (string)Session["user_id"];
      if (string.IsNullOrEmpty(userId))
      {
        TempData["message"] = "Please sign in to continue";
        return Redirect("/");
      }

      OAuthHelper oauthHelper = new OAuthHelper(authority, appId, appSecret);
      string redirectUri = Url.Action("Authorize", "Home", null, Request.Url.Scheme);

      string accessToken = await oauthHelper.GetAccessToken(userId, redirectUri);
      if (string.IsNullOrEmpty(accessToken))
      {
        TempData["message"] = "Please sign in to continue";
        return Redirect("/");
      }

      var client = new Outlook();
      client.anchorMailbox = (string)Session["user_name"];
      ViewBag.UserName = client.anchorMailbox;

      DateTime viewStart = DateTime.Now.ToUniversalTime();
      DateTime viewEnd = viewStart.AddHours(3);
      var result = await client.GetCalendarView(accessToken, client.anchorMailbox, viewStart, viewEnd);

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