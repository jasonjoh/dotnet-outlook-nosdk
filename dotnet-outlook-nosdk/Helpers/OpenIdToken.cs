using Newtonsoft.Json;
using System;
using System.Globalization;
using System.Text;

namespace dotnet_outlook_nosdk.Helpers
{
  public class OpenIdToken
  {
    public string aud;
    public string iss;
    public string iat;
    public string nbf;
    public string exp;
    public string ver;
    public string tid;
    public string oid;
    public string email;
    public string sub;
    public string name;
    public string nonce;
    public string auth_time;

    public static OpenIdToken ParseOpenIdToken(string id_token)
    {
      string encodedOpenIdToken = id_token;

      string decodedToken = Base64UrlDecodeJwtTokenPayload(encodedOpenIdToken);

      OpenIdToken token = JsonConvert.DeserializeObject<OpenIdToken>(decodedToken);

      return token;
    }

    public static bool ValidateOpenIdToken(string id_token, string nonce)
    {
      if (String.IsNullOrEmpty(nonce))
      { // nothing to validate
        return false;
      }

      OpenIdToken token = ParseOpenIdToken(id_token);
      if (token.nonce.Equals(nonce))
      { // cheap validation. must add signature validation.
        return true;
      }

      return false;
    }

    private static string Base64UrlDecodeJwtTokenPayload(string base64UrlEncodedJwtToken)
    {
      string payload = base64UrlEncodedJwtToken.Split('.')[1];
      return Base64UrlEncoder.Decode(payload);
    }
  }

  // From: Jason Johnston@https://github.com/jasonjoh/office365-azure-guides/blob/master/code/parse-token.cs
  public static class Base64UrlEncoder
  {
    static char Base64PadCharacter = '=';
    static string DoubleBase64PadCharacter = String.Format(CultureInfo.InvariantCulture, "{0}{0}", Base64PadCharacter);
    static char Base64Character62 = '+';
    static char Base64Character63 = '/';
    static char Base64UrlCharacter62 = '-';
    static char Base64UrlCharacter63 = '_';

    public static byte[] DecodeBytes(string arg)
    {
      string s = arg;
      s = s.Replace(Base64UrlCharacter62, Base64Character62); // 62nd char of encoding
      s = s.Replace(Base64UrlCharacter63, Base64Character63); // 63rd char of encoding
      switch (s.Length % 4) // Pad 
      {
        case 0:
          break; // No pad chars in this case
        case 2:
          s += DoubleBase64PadCharacter; break; // Two pad chars
        case 3:
          s += Base64PadCharacter; break; // One pad char
        default:
          throw new ArgumentException("Illegal base64url string!", arg);
      }
      return Convert.FromBase64String(s); // Standard base64 decoder
    }

    public static string Decode(string arg)
    {
      return Encoding.UTF8.GetString(DecodeBytes(arg));
    }
  }
}