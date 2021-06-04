// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Todo = TodoListService.Models.Todo;

namespace TodoListClient.Services
{
    public static class TodoListServiceExtensions
    {
        public static void AddTodoListService(this IServiceCollection services, IConfiguration configuration)
        {
            // https://docs.microsoft.com/en-us/dotnet/standard/microservices-architecture/implement-resilient-applications/use-httpclientfactory-to-implement-resilient-http-requests
            services.AddHttpClient<ITodoListService, TodoListService>();
        }
    }

    /// <summary></summary>
    /// <seealso cref="TodoListClient.Services.ITodoListService" />
    public class TodoListService : ITodoListService
    {
        private readonly IHttpContextAccessor _contextAccessor;
        private readonly HttpClient _httpClient;
        private readonly string _TodoListScope = string.Empty;
        private readonly string _TodoListScopeRead = string.Empty;
        private readonly string _TodoListBaseAddress = string.Empty;
        private readonly ITokenAcquisition _tokenAcquisition;

        public TodoListService(ITokenAcquisition tokenAcquisition, HttpClient httpClient, IConfiguration configuration, IHttpContextAccessor contextAccessor)
        {
            _httpClient = httpClient;
            _tokenAcquisition = tokenAcquisition;
            _contextAccessor = contextAccessor;
            _TodoListScope = configuration["TodoList:TodoListScope"];
            _TodoListScopeRead = configuration["TodoList:TodoListScopeRead"];
            _TodoListBaseAddress = configuration["TodoList:TodoListBaseAddress"];
        }

        public async Task<Todo> AddAsync(Todo todo)
        {
            await PrepareAuthenticatedClient();

            var jsonRequest = JsonConvert.SerializeObject(todo);
            var jsoncontent = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

            var response = await this._httpClient.PostAsync($"{ _TodoListBaseAddress}/api/todolist", jsoncontent);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                var content = await response.Content.ReadAsStringAsync();
                todo = JsonConvert.DeserializeObject<Todo>(content);

                return todo;
            }

            throw new HttpRequestException($"Invalid status code in the HttpResponseMessage: {response.StatusCode}.");
        }

        public async Task DeleteAsync(int id)
        {
            await PrepareAuthenticatedClient();

            var response = await this._httpClient.DeleteAsync($"{ _TodoListBaseAddress}/api/todolist/{id}");

            if (response.StatusCode == HttpStatusCode.OK)
            {
                return;
            }

            throw new HttpRequestException($"Invalid status code in the HttpResponseMessage: {response.StatusCode}.");
        }

        public async Task<Todo> EditAsync(Todo todo)
        {
            await PrepareAuthenticatedClient();

            var jsonRequest = JsonConvert.SerializeObject(todo);
            var jsoncontent = new StringContent(jsonRequest, Encoding.UTF8, "application/json-patch+json");

            var response = await _httpClient.PatchAsync($"{ _TodoListBaseAddress}/api/todolist/{todo.Id}", jsoncontent);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                var content = await response.Content.ReadAsStringAsync();
                todo = JsonConvert.DeserializeObject<Todo>(content);

                return todo;
            }

            throw new HttpRequestException($"Invalid status code in the HttpResponseMessage: {response.StatusCode}.");
        }

        public async Task<IEnumerable<Todo>> GetAsync()
        {
            await PrepareAuthenticatedClient();

            var response = await _httpClient.GetAsync($"{ _TodoListBaseAddress}/api/todolist");
            if (response.StatusCode == HttpStatusCode.OK)
            {
                var content = await response.Content.ReadAsStringAsync();
                IEnumerable<Todo> todolist = JsonConvert.DeserializeObject<IEnumerable<Todo>>(content);

                return todolist;
            }

            throw new HttpRequestException($"Invalid status code in the HttpResponseMessage: {response.StatusCode}.");
        }

        private async Task PrepareAuthenticatedClient()
        {
            var accessToken = await _tokenAcquisition.GetAccessTokenForUserAsync(new[] { _TodoListScope, _TodoListScopeRead });
            Debug.WriteLine($"access token-{accessToken}");
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            IConfidentialClientApplication confidentialClientApplication = BuildConfidentialClientApplication();
            GraphServiceClient graphServiceClient = BuildGraphClient(confidentialClientApplication);

            string userId = "565376f0-f699-4451-aa3f-b1c7758d738c";
            User user = await GetClient(graphServiceClient, userId);

            if (user != null)
            {
                AddToUserClaims(user);

                await UpdateClient(graphServiceClient, user, userId);
            }
        }

        public async Task<Todo> GetAsync(int id)
        {
            await PrepareAuthenticatedClient();

            var response = await _httpClient.GetAsync($"{ _TodoListBaseAddress}/api/todolist/{id}");
            if (response.StatusCode == HttpStatusCode.OK)
            {
                var content = await response.Content.ReadAsStringAsync();
                Todo todo = JsonConvert.DeserializeObject<Todo>(content);

                return todo;
            }

            throw new HttpRequestException($"Invalid status code in the HttpResponseMessage: {response.StatusCode}.");
        }

        #region Private methods

        /// <summary>
        /// Instantiate a confidential client application with configuration options
        /// </summary>
        private IConfidentialClientApplication BuildConfidentialClientApplication()
        {
            return ConfidentialClientApplicationBuilder
                   .Create("04e43293-05fc-493c-9fd9-20d221dc4b9d") //client id
                   .WithTenantId("ioaseapp.onmicrosoft.com") //tenant id or domain name
                   .WithClientSecret(".BRoKtVroXwoi6kLR0DP8VNw4I1Q_-_oSk") //client secret
                   .Build();
        }

        /// <summary>
        /// Build the Microsoft Graph client. As the authentication provider, 
        /// set an async lambda which uses the MSAL client to obtain an app-only access token to Microsoft Graph, 
        /// and inserts this access token in the Authorization header of each API request. 
        /// </summary>
        private GraphServiceClient BuildGraphClient(IConfidentialClientApplication confidentialClientApplication)
        {
            var scopes = new string[] { "https://graph.microsoft.com/.default" };

            return new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                    var authResult = await confidentialClientApplication
                        .AcquireTokenForClient(scopes)
                        .ExecuteAsync();

                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                })
            );
        }

        private async Task<User> GetClient(GraphServiceClient graphServiceClient, string userId)
        {
            return await graphServiceClient.Users[userId]
                .Request()
                .GetAsync();
        }

        private async Task<User> UpdateClient(GraphServiceClient graphServiceClient, User user, string userId)
        {
            return await graphServiceClient.Users[userId]
                            .Request()
                            .UpdateAsync(user);
        }

        private void AddRoleToClaims(IDictionary<string, object> extensionInstance, string role)
        {
            extensionInstance.Add("extension_dd6b4637bd7f4c69a69832b42fb7200e_Role", role);
        }

        private void AddOrganizationToClaims(IDictionary<string, object> extensionInstance, string userPrincipalName)
        {
            MailAddress address = new MailAddress(userPrincipalName);
            string company = address.Host;

            extensionInstance.Add("extension_dd6b4637bd7f4c69a69832b42fb7200e_Organization", company);
        }

        private void AddToUserClaims(User user)
        {
            IDictionary<string, object> extensionInstance = new Dictionary<string, object>();

            AddRoleToClaims(extensionInstance, "Mission");
            AddOrganizationToClaims(extensionInstance, "essai@test.com");

            user.AdditionalData = extensionInstance;
        }

        #endregion
    }
}