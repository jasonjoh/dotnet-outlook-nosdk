# Outlook REST API SDK-less Sample

An example of calling the Outlook API without using the NuGet libraries. The NuGet libraries support the v2.0 API endpoint (and the older library supports the v1.0 endpoint) but you can't use the beta endpoint. By implementing the REST calls yourself, you can call whichever endpoint you want.

## Running the sample

This sample was created in Visual Studio 2015 Update 1, using the **ASP.NET Web Application** template, with the following details:

- Specific template: **MVC** under **ASP.NET 4.5.2 Templates**
- Authentication: No Authentication

I have not tested this project with Visual Studio 2013.

1. Download, clone, or fork this repository.
1. Open the `dotnet-outlook-nosdk.sln` file in Visual Studio 2015.
1. In **Solution Explorer**, right click the **Solution 'dotnet-outlook-sdk'** node and choose **Restore NuGet Packages**.
1. Create an app registration for the sample at https://apps.dev.microsoft.com to get an app ID and secret. For instructions on doing this, see section 3 of [this tutorial](https://dev.outlook.com/RestGettingStarted/Tutorial/dotnet). The redirect URIs for running this sample on your development machine are:

        http://localhost:34301/Home/Authorize
        https://localhost:44300/Home/Authorize

1. Open the `Web.config` file and locate the following lines:

        <add key="ida:ClientID" value="YOUR APP ID HERE" />
        <add key="ida:ClientSecret" value="YOUR APP SECRET HERE" />
        
  Replace the values with the app ID and app secret generated in the previous step.

1. Press **F5** to run the sample.

## Copyright ##

Copyright (c) Microsoft. All rights reserved.

----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Outlook Dev Blog](http://blogs.msdn.com/b/exchangedev/)
