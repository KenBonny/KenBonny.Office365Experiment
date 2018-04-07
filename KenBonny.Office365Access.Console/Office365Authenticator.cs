using System.Security.Authentication;
using RestSharp;
using RestSharp.Authenticators;

namespace KenBonny.Office365Access.Console
{
    internal class Office365Authenticator : IAuthenticator
    {
        private readonly IRestClient _authenticationEndpoint;
        private readonly IRestRequest _authenticationRequest;
        private Office365AuthenticationToken _office365AuthenticationResponse;

        public Office365Authenticator(string clientId, string clientSecret)
        {
            _authenticationEndpoint = CreateAuthenticationEndpoint();
            _authenticationRequest = CreateAuthenticationRequest(clientId, clientSecret);
            _office365AuthenticationResponse = new Office365AuthenticationToken();
        }

        public void Authenticate(IRestClient client, IRestRequest request)
        {
            if (_office365AuthenticationResponse.IsExpired)
            {
                FetchNewToken();
            }

            request.AddHeader("Authorization", _office365AuthenticationResponse.Token);
        }

        private void FetchNewToken()
        {
            var response = _authenticationEndpoint.Execute<Office365AuthenticationToken>(_authenticationRequest);
            if (response.IsSuccessful && response.Data != null)
            {
                _office365AuthenticationResponse = response.Data;
            }
            else
            {
                if (response.ErrorException != null)
                    throw response.ErrorException;
                
                if (!string.IsNullOrWhiteSpace(response.ErrorMessage))
                    throw new AuthenticationException(response.ErrorMessage);
                
                if (!string.IsNullOrWhiteSpace(response.Content))
                    throw new AuthenticationException(response.Content);
                
                throw new AuthenticationException("Could not authenticate");
            }
        }

        private static IRestClient CreateAuthenticationEndpoint()
        {
            const string authenticationUrl = "https://login.microsoftonline.com/";
            return new RestClient(authenticationUrl);
        }

        private static IRestRequest CreateAuthenticationRequest(string clientId, string clientSecret)
        {
            var request = new RestRequest("/24b1c222-db0f-4bef-b980-e08ac1762707/oauth2/v2.0/token", Method.POST);
            request.AddParameter("grant_type", "client_credentials");
            request.AddParameter("client_id", clientId);
            request.AddParameter("client_secret", clientSecret);
            request.AddParameter("scope", "https://graph.microsoft.com/.default");
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            return request;
        }
    }
}