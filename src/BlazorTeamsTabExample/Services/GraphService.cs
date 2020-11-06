using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace BlazorTeamsTabExample.Services
{
    public class GraphService
    {
        public async Task<String> GetMeWithHttpClient(string token)
        {
            var _httpclient = new HttpClient();
            _httpclient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("bearer", token);
            var response = await _httpclient.GetAsync("https://graph.microsoft.com/v1.0/me/");

            return response.Content.ReadAsStringAsync().Result;
        }

        public async Task<List<T>> GetAllEntities<T>(string entity, string token)
        {
            var entities = new List<T>();

            var query = $"https://graph.microsoft.com/v1.0/{entity}";

            var hasNextPage = true;

            while (hasNextPage)
            {
                var _httpclient = new HttpClient();

                var request = new HttpRequestMessage(HttpMethod.Get, query);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                try
                {
                    var response = await _httpclient.SendAsync(request);
                    var json = await response.Content.ReadAsStringAsync();
                    var jObject = JObject.Parse(json);
                    var jObjects = (JArray)jObject["value"];

                    entities.AddRange(jObjects.ToObject<List<T>>());

                    query = (string)jObject["@odata.nextLink"];

                    if (string.IsNullOrEmpty(query))
                    {
                        hasNextPage = false;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    hasNextPage = false;
                }
            }

            return entities;
        }
    }
}
