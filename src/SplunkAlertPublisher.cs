using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SplunkMailProcessor
{
    public class SplunkAlertPublisher
    {
        HttpClient client;
        string url;

        public SplunkAlertPublisher(string url, string ApiKey)
        {
            client = new HttpClient();
            this.url = url;
            //client.BaseAddress = new Uri(url);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("Authorization", "Splunk " + ApiKey);
        }

        public async Task<bool> CreateAlertAsync(Alert alert)
        {
            var content = new StringContent(JsonConvert.SerializeObject(alert), Encoding.UTF8, "application/json");
            HttpResponseMessage response = await client.PostAsync(
                url, content);
            response.EnsureSuccessStatusCode();

            // return URI of the created resource.
            return true;
        }
    }
}
