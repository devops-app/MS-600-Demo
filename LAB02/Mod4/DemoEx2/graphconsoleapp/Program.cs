using System;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
    public class Program
    {
        private static GraphServiceClient? _graphClient;

        private static IConfigurationRoot? LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                                  .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                                  .AddJsonFile("appsettings.json", false, true)
                                  .Build();

                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithRedirectUri(redirectUri)
                                                    .WithClientSecret(clientSecret)
                                                    .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }
        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphClient = new GraphServiceClient(authenticationProvider);
            return _graphClient;
        }


        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var config = LoadAppSettings();
            if (config == null)
            {
                Console.WriteLine("Invalid appsettings.json file.");
                return;
            }
            var client = GetAuthenticatedGraphClient(config);

            // request 1 - current user's photo

            var requestUserPhoto = client.Me.Photo.Request();
            var resultsUserPhoto = requestUserPhoto.GetAsync().Result;
            // display photo metadata
            Console.WriteLine("                Id: " + resultsUserPhoto.Id);
            Console.WriteLine("media content type: " + resultsUserPhoto.AdditionalData["@odata.mediaContentType"]);
            Console.WriteLine("        media etag: " + resultsUserPhoto.AdditionalData["@odata.mediaEtag"]);

            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(requestUserPhoto.GetHttpRequestMessage().RequestUri);

            // get actual photo
            var requestUserPhotoFile = client.Me.Photo.Content.Request();
            var resultUserPhotoFile = requestUserPhotoFile.GetAsync().Result;

            // create the file
            var profilePhotoPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "profilePhoto_" + resultsUserPhoto.Id + ".jpg");
            var profilePhotoFile = System.IO.File.Create(profilePhotoPath);
            resultUserPhotoFile.Seek(0, System.IO.SeekOrigin.Begin);
            resultUserPhotoFile.CopyTo(profilePhotoFile);
            Console.WriteLine("Saved file to: " + profilePhotoPath);

            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(requestUserPhoto.GetHttpRequestMessage().RequestUri);

            /*
            // request 2 - user's manager
            var userId = "{{REPLACE_WITH_USER_ID}}";
            var requestUserManager = client.Users[userId]
                                            .Manager
                                            .Request();
            var resultsUserManager = requestUserManager.GetAsync().Result;
            Console.WriteLine("   User: " + userId);
            Console.WriteLine("Manager: " + resultsUserManager.Id);
            var resultsUserManagerUser = resultsUserManager as Microsoft.Graph.User;
            if (resultsUserManagerUser != null)
            {
            Console.WriteLine("Manager: " + resultsUserManagerUser.DisplayName);
            Console.WriteLine(resultsUserManager.Id + ": " + resultsUserManagerUser.DisplayName + " <" + resultsUserManagerUser.Mail + ">");
            }

            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(requestUserManager.GetHttpRequestMessage().RequestUri);
            */
        }
    }
}