using System;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;
using System.Threading.Tasks;

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

        private static async Task<Microsoft.Graph.Group> CreateGroupAsync(GraphServiceClient client)
        {
            // create object to define members & owners as 'additionalData'
            var additionalData = new Dictionary<string, object>();
            additionalData.Add("owners@odata.bind",
              new string[] {
    "https://graph.microsoft.com/v1.0/users/d280a087-e05b-4c23-b073-738cdb82b25e"
              }
            );
            additionalData.Add("members@odata.bind",
              new string[] {
    "https://graph.microsoft.com/v1.0/users/70c095fe-df9d-4250-867d-f298e237d681",
    "https://graph.microsoft.com/v1.0/users/8c2da469-1eba-47a4-9322-ee0ddd24d99a"
              }
            );
            var group = new Microsoft.Graph.Group
            {
                AdditionalData = additionalData,
                Description = "My first group created with the Microsoft Graph .NET SDK",
                DisplayName = "My First Group",
                GroupTypes = new List<String>() { "Unified" },
                MailEnabled = true,
                MailNickname = "myfirstgroup01",
                SecurityEnabled = false
            };
            var requestNewGroup = client.Groups.Request();
            return await requestNewGroup.AddAsync(group);
        }
        private static async Task<Microsoft.Graph.Team> TeamifyGroupAsync(GraphServiceClient client, string groupId)
        {
            var team = new Microsoft.Graph.Team
            {
                MemberSettings = new TeamMemberSettings
                {
                    AllowCreateUpdateChannels = true,
                    ODataType = null
                },
                MessagingSettings = new TeamMessagingSettings
                {
                    AllowUserEditMessages = true,
                    AllowUserDeleteMessages = true,
                    ODataType = null
                },
                ODataType = null
            };

            var requestTeamifiedGroup = client.Groups[groupId].Team.Request();
            return await requestTeamifiedGroup.PutAsync(team);
        }
        private static async Task DeleteTeamAsync(GraphServiceClient client, string groupIdToDelete)
        {
            await client.Groups[groupIdToDelete].Request().DeleteAsync();
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

            var profileResponse = client.Me.Request().GetAsync().Result;
            Console.WriteLine("Hello " + profileResponse.DisplayName);

            // request 1 - create new group
            Console.WriteLine("\n\nREQUEST 1 - CREATE A GROUP:");
            var requestNewGroup = CreateGroupAsync(client);
            requestNewGroup.Wait();
            Console.WriteLine("New group ID: " + requestNewGroup.Id);

            /*
            // request 2 - teamify group
            // get new group ID
            var requestGroup = client.Groups.Request()
                                            .Select("Id")
                                            .Filter("MailNickname eq 'myfirstgroup01'");
            var resultGroup = requestGroup.GetAsync().Result;

            // teamify group
            var teamifiedGroup = TeamifyGroupAsync(client, resultGroup[0].Id);
            teamifiedGroup.Wait();
            Console.WriteLine(teamifiedGroup.Result.Id);

            // request 3: delete group
            var deleteTask = DeleteTeamAsync(client, resultGroup[0].Id);
            deleteTask.Wait();
            */
        }
    }
}