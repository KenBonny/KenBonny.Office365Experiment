using System;
using Newtonsoft.Json;

namespace KenBonny.Office365Access.Console
{
    internal class Office365AuthenticationToken
    {

        [JsonProperty("access_token")]
        public string AccessToken { get; set; }

        [JsonProperty("token_type")]
        public string TokenType { get; set; }

        [JsonProperty("expires_in")]
        public int ExpiresIn { get; set; }

        [JsonProperty("expires_on")]
        public long ExpiresOn { get; set; }

        [JsonProperty("resource")]
        public string Resource { get; set; }

        public string Token => $"{TokenType} {AccessToken}";

        public bool IsExpired => DateTime.Now > ExpirationDate;

        private DateTime ExpirationDate => CreateExpirationDate();

        private DateTime CreateExpirationDate()
        {
            var expireDate = DateTime.MinValue;
            if (ExpiresOn > 0)
            {
                expireDate = new DateTime(ExpiresOn);
            }
            else if (ExpiresIn > 0)
            {
                expireDate = DateTime.Now.AddSeconds(ExpiresIn);
            }
            return expireDate;
        }
    }
}