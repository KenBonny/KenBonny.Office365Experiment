using System;
using System.Collections.Generic;
using RestSharp;

namespace KenBonny.Office365Access.Console
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            System.Console.WriteLine("//--- Outlook mail fetcher ---\\\\");

            if (args.Length != 2)
            {
                System.Console.WriteLine("Please pass the clientId and clientSecret as arguements in that order");
                return;
            }

            try
            {
                var clientId = args[0];
                var clientSecret = args[1];

                var emails = GetEmails(clientId, clientSecret);
                foreach (var email in emails)
                {
                    var isReadPlaceholder = email.IsRead ? string.Empty : "*";
                    System.Console.WriteLine($" - {email.Subject} {isReadPlaceholder} ({email.From.EmailAddress.Name})");
                }
            }
            catch (Exception exception)
            {
                System.Console.WriteLine(exception);
            }
        }

        private static IReadOnlyCollection<Email> GetEmails(string clientId, string clientSecret)
        {
            var officeClient = new RestClient("https://graph.microsoft.com/v1.0/me")
            {
                Authenticator = new Office365Authenticator(clientId, clientSecret)
            };
            var request = new RestRequest("messages", Method.GET);
            var response = officeClient.Execute<EmailCollection>(request);
            return response.IsSuccessful ? response.Data.Emails : new Email[0];
        }
    }
}