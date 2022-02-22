using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace O365.Mail
{
    public class AuthenticationProvider : IAuthenticationProvider
    {
        private readonly IConfidentialClientApplication _clientApplication;
        private readonly string[] _scopes;

        public AuthenticationProvider(IConfidentialClientApplication clientApplication, string[] scopes)
        {
            _clientApplication = clientApplication;
            _scopes = scopes;
        }

        // Getting authorization using the Access Token
        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            string token = await GetTokenAsync();
            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
        }

        // Getting access Token
        public async Task<string> GetTokenAsync()
        {
            AuthenticationResult authResult = null;
            authResult = await _clientApplication.AcquireTokenForClient(_scopes).ExecuteAsync();

            return authResult.AccessToken;
        }
    }
}
