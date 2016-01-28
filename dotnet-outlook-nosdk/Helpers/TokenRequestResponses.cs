namespace dotnet_outlook_nosdk.Helpers
{
  public class TokenRequestSuccessResponse
  {
    public string token_type;
    public string expires_in;
    public string expires_on;
    public string not_before;
    public string resource;
    public string access_token;
    public string refresh_token;
    public string id_token;
    public string scope;
    public string pwd_exp;
    public string pwd_url;
  }

  public class TokenRequestErrorResponse
  {
    //{
    //  "error": "invalid_client",
    //  "error_description": "AADSTS70002: Error ...",
    //  "error_codes": [
    //    70002,
    //    50012
    //  ],
    //  "timestamp": "2015-02-07 18:44:09Z",
    //  "trace_id": "dabcfa26-ea8d-46c5-81bc-ff57a0895629",
    //  "correlation_id": "8e270f2d-ba05-42fb-a7ab-e819d142c843",
    //  "submit_url": null,
    //  "context": null
    //}
    public string error;
    public string error_description;
    public string[] error_codes;
    public string timestamp;
    public string trace_id;
    public string correlation_id;
    public string submit_url;
    public string context;
  }
}