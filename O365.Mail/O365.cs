using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace O365.Mail
{
    public class O365 : IO365
    {
        private static GraphServiceClient _client;
        private readonly IConfidentialClientApplication _cca;

        public O365(string tenantId, string clientId, string clientSecret)
        {
            try
            {
                // Authentication to Azure AD
                _cca = ConfidentialClientApplicationBuilder.Create(clientId)
                    .WithTenantId(tenantId)
                    .WithClientSecret(clientSecret)
                    .Build();

                // Default scopes are sufficient
                List<string> scopes = new List<string>()
                {
                    "https://graph.microsoft.com/.default"
                };

                var authProvider = new AuthenticationProvider(_cca, scopes.ToArray());
                _client = new GraphServiceClient(authProvider);
            }

            catch (Exception ex)
            {
                throw new Exception($"Error connecting the API, please check permissions.\n {ex.Message}", ex);
            }
        }

        public async Task SendMail(Message message)
        {
            try
            {
                string sendinAdress = message.From.EmailAddress.Address;
                await _client.Users[sendinAdress].SendMail(message, false).Request().PostAsync();
            }

            catch(Exception ex)
            {
                throw new Exception($"Error while sending email.\n {ex.Message}", ex);
            }
        }
    }
}
