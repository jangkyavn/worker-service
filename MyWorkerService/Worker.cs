using Microsoft.AspNetCore.SignalR.Client;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace MyWorkerService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private HubConnection connection;
        private string _accessToken;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        public override Task StartAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service Started at: {time}", DateTimeOffset.Now);

            RunAsync().Wait();

            connection = new HubConnectionBuilder()
                .WithUrl("https://localhost:44383/notify", options =>
                {
                    options.AccessTokenProvider = () => Task.FromResult(_accessToken);
                })
                .Build();

            return base.StartAsync(cancellationToken);
        }

        private async Task RunAsync()
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://localhost:44383/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                HttpResponseMessage response = await client.PostAsJsonAsync("api/Auth/login", new
                {
                    userName = "admin",
                    password = "123456"
                });

                if (response.IsSuccessStatusCode)
                {
                    LoginReponse loginReponse = await response.Content.ReadAsAsync<LoginReponse>();
                    _accessToken = loginReponse.Token;

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
                    HttpResponseMessage productResponse = await client.GetAsync("api/Product");
                    if (productResponse.IsSuccessStatusCode)
                    {
                        Product[] products = await productResponse.Content.ReadAsAsync<Product[]>();

                        _logger.LogInformation("products: {0}", JsonConvert.SerializeObject(products));
                    }
                }
            }
        }

        public override Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service Stopped at: {time}", DateTimeOffset.Now);
            return base.StopAsync(cancellationToken);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            //while (!stoppingToken.IsCancellationRequested)
            //{
            //    //if (GetResponse())
            //    //{
            //    //    _logger.LogInformation("Response is OK: {time}", DateTimeOffset.Now);
            //    //}
            //    //else
            //    //{
            //    //    _logger.LogInformation("No Response: {time}", DateTimeOffset.Now);
            //    //}    

            //    await Task.Delay(0, stoppingToken);
            //}

            connection.On<string>("ReceiveMessage1", (message) =>
            {
                _logger.LogInformation("Message: {0}", message);
            });

            await connection.StartAsync();
        }

        public bool GetResponse()
        {
            return true;
        }
    }

    public class Product
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
    }

    public class LoginReponse
    {
        public string Token { get; set; }
        public bool Status { get; set; }
        public string Message { get; set; }
    }
}
