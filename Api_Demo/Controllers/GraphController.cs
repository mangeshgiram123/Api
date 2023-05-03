using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text;
using Microsoft.Extensions.Logging;
using System.Net.Mail;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;


namespace IntegratedEmail_CalendarEvent_GraphAPI_Core.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GraphController : ControllerBase
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly IConfiguration _config;
        public GraphController(IHttpClientFactory httpClientFactory, IConfiguration config)
        {
            _httpClientFactory = httpClientFactory;
            _config = config; //New

        }

        [Route("Email")]
        [HttpGet]
        public async Task<IActionResult> Get()
        {
            var azureAdSection = _config.GetSection("AzureAd"); //New

            string instance = azureAdSection.GetValue<string>("Instance");
            string tenantId = azureAdSection.GetValue<string>("TenantId");
            string clientId = azureAdSection.GetValue<string>("ClientId");
            string clientSecret = azureAdSection.GetValue<string>("ClientSecret");

            string tokenEndpoint = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            using (HttpClient client = _httpClientFactory.CreateClient())
            {
                var requestContent = new FormUrlEncodedContent(new[]
                {

          new KeyValuePair <string, string>("grant_type", "client_credentials"),
          new KeyValuePair<string, string>("client_id", clientId),
          new KeyValuePair<string, string>("client_secret", clientSecret),
          new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
        });

                HttpResponseMessage response = await client.PostAsync(tokenEndpoint, requestContent);
                string responseContent = await response.Content.ReadAsStringAsync();
                var tokenResponse = JsonSerializer.Deserialize<JsonElement>(responseContent);

                var accessToken = tokenResponse.GetProperty("access_token").GetString();

                string mailUser = "navjot.thakur@yashtechnologies841.onmicrosoft.com";
                string sendMailEndpoint = $"https://graph.microsoft.com/v1.0/users/{mailUser}/sendMail";

                var message = new Dictionary<string, object>()

        {

          {"message", new Dictionary<string, object>()
          {
              { "subject", "Test Email using the Graph API" },
              { "body", new Dictionary<string, object>()
              {
                { "contentType", "Text" },
                { "content", "Hello! This is a test email for MS AZURE - Graph API - Demo for Send Email Functionality!" }
              }
              },

              { "toRecipients",  new object[] {
              new Dictionary<string, object>()
              {
                { "emailAddress", new Dictionary<string, object>()
                {
                 { "address", "ashupardhi00u@gmail.com" }
                }
                }
              }
             }},
            }
          },
          {"saveToSentItems", "true" }
        };

                var jsonMessage = JsonSerializer.Serialize(message);
                var content = new StringContent(jsonMessage, Encoding.UTF8, "application/json");
                var request = new HttpRequestMessage(HttpMethod.Post, sendMailEndpoint);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                request.Content = content;

                HttpResponseMessage sendMailResponse = await client.SendAsync(request);
                string SendMailResponseContent = await sendMailResponse.Content.ReadAsStringAsync();

                if (sendMailResponse.IsSuccessStatusCode)
                {
                    return Ok("Email sent successfully!");
                }

                else
                {
                    return BadRequest("Failed to send the email!");
                }

            }

        }

        [Route("MeetingLink")]
        [HttpPost]
        public async Task<IActionResult> Post()
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var azureAdSection = _config.GetSection("AzureAd"); //New

            string instance = azureAdSection.GetValue<string>("Instance");
            string tenantId = azureAdSection.GetValue<string>("TenantId");
            string clientId = azureAdSection.GetValue<string>("ClientId");
            string clientSecret = azureAdSection.GetValue<string>("ClientSecret");

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            var userPrincipalName = "navjot.thakur@yashtechnologies841.onmicrosoft.com";
            var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            var requestBody = new Event
            {
                Subject = "Let's go for VS Migration of the Legacy Apps!",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "Does noon work for you?",
                },

                Start = new DateTimeTimeZone
                {
                    DateTime = "2023-05-03T10:00:00",
                    TimeZone = "Pacific Standard Time",
                },

                End = new DateTimeTimeZone
                {
                    DateTime = "2023-05-03T11:00:00",
                    TimeZone = "Pacific Standard Time",
                },

                Location = new Location
                {
                    DisplayName = "Polaris Meeting",
                },

                Attendees = new List<Attendee>
        {
          new Attendee
          {
            EmailAddress = new EmailAddress
            {
              Address = "ashupardhi00u@gmail.com",
              Name = "Navjot Thakur",
            },
            Type = AttendeeType.Required,
          },
        },

                AllowNewTimeProposals = true,
                IsOnlineMeeting = true,
                OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness,
            };

            var result = await graphClient.Users[userPrincipalName].Events.PostAsync(requestBody, (requestConfiguration) =>
            {
                requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Pacific Standard Time\"");
            });
            return Ok(result);
        }
    }
}