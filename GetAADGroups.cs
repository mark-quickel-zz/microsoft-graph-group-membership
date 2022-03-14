using System;
using System.Collections;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Azure.Core;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Formatters;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Newtonsoft.Json;

namespace Microsoft.EE
{
    public class GetAADGroups
    {
        [FunctionName("GetAADGroups")]
        [OpenApiOperation(operationId: "Run", tags: new[] { "name" })]
        [OpenApiSecurity("function_key", SecuritySchemeType.ApiKey, Name = "code", In = OpenApiSecurityLocationType.Query)]
        [OpenApiParameter(name: "name", In = ParameterLocation.Query, Required = true, Type = typeof(string), Description = "The **Name** parameter")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "text/plain", bodyType: typeof(string), Description = "The OK response")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name; //unused presently, can be passed to the Graph API to query specific names

            log.LogInformation("Invocation of GetAccessToken()");
            var token = GetAccessToken();
            if (String.IsNullOrWhiteSpace(token)) return new UnauthorizedResult();
            log.LogInformation("Invocation of GetAADGroupsFromGraph()");
            return new OkObjectResult(GetAADGroupsFromGraph(token));
        }

        private IEnumerable<GroupEntity> GetAADGroupsFromGraph(string token)
        {
            var uri = new Uri("https://graph.microsoft.com/v1.0/groups");

            var client = new WebClient();
            client.Headers.Add("Accept", "application/json");
            client.Headers.Add("Content-Type", "application/json; charset=utf-8");
            client.Headers.Add("Authorization", $"Bearer {token}");
            var result = client.DownloadString(uri);

            var groupsEntity = JsonConvert.DeserializeObject<GraphGroupCollectionEntity>(result);
            return groupsEntity.value.Where(x => x.mailEnabled == true).OrderBy(x => x.displayName).Take(50);

        }

        private void wc_GetAADGroupsFromGraphCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            string text = e.Result;
        }

        private string GetAccessToken()
        {
            var scope = "https://graph.microsoft.com/";
            var identityEndpoint = Environment.GetEnvironmentVariable("IDENTITY_ENDPOINT");
            var identityHeader = Environment.GetEnvironmentVariable("IDENTITY_HEADER");
            var tokenAuthUri = $"{identityEndpoint}?resource={scope}&api-version=2019-08-01";

            var uri = new Uri(tokenAuthUri);
            var client = new WebClient();
            client.Headers.Add("X-IDENTITY-HEADER", identityHeader);
            var result = client.DownloadString(uri);

            var tokenEntity = JsonConvert.DeserializeObject<TokenEntity>(result);
            return tokenEntity.access_token;
        }

    }

    internal class TokenEntity
    {
        public string access_token { get; set; }
    }

    internal class GraphGroupCollectionEntity
    { 
        public List<GroupEntity> value { get; set; }
    }

    internal class GroupEntity
    {
        public Guid id { get; set; }
        public string displayName { get; set; }
        public string mail { get; set; }
        public bool mailEnabled { get; set; }

    }
}

