using Microsoft.AspNetCore.Mvc;
using MongoDB.Bson;
using MongoDB.Driver;
using ClosedXML.Excel;

namespace CrudPet.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ObakController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<ObakController> _logger;
        //private readonly IHttpClientFactory _httpClientFactory;

        public ObakController(
            ILogger<ObakController> logger 
            //IHttpClientFactory httpClientFactory
            )
        {
            _logger = logger;
            //_httpClientFactory = httpClientFactory;
        }

        //[HttpGet("GetWeatherForecast")]
        //public IEnumerable<WeatherForecast> Get()
        //{
        //    return Enumerable.Range(1, 5).Select(index => new WeatherForecast
        //    {
        //        Date = DateTime.Now.AddDays(index),
        //        TemperatureC = Random.Shared.Next(-20, 55),
        //        Summary = Summaries[Random.Shared.Next(Summaries.Length)]
        //    })
        //    .ToArray();
        //}
        [HttpGet("GetData")]
        public async Task<IActionResult> GetData()
        {
            const string connectionUri = "mongodb+srv://anubis1080p:vHNelDfzCdi0eR3W@obak.iyvad.mongodb.net/?retryWrites=true&w=majority&appName=OBAK";
            var settings = MongoClientSettings.FromConnectionString(connectionUri);
            // Set the ServerApi field of the settings object to set the version of the Stable API on the client
            settings.ServerApi = new ServerApi(ServerApiVersion.V1);
            // Create a new client and connect to the server
            var client = new MongoClient(settings);
            // Send a ping to confirm a successful connection
            try
            {
                string filePath = @"H:\Projects\OBAK\Obak.xlsx";
                //Stream xlsxStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using (Stream xlsxStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {

                    var workbook = new XLWorkbook(xlsxStream);
                    var worksheet = workbook.Worksheet(1); // Assuming the data is in the first worksheet

                    // Get the first row as headers
                    var headerRow = worksheet.Row(1);
                    var headers = new List<string>();

                    // Store column headers (property names)
                    foreach (var headerCell in headerRow.CellsUsed())
                    {
                        headers.Add(headerCell.Value.ToString());
                    }

                    // List to hold the MongoDB documents
                    var documents = new List<BsonDocument>();

                    // Loop through each data row (starting from the second row)
                    foreach (var row in worksheet.RowsUsed().Skip(1)) // Skip the header row
                    {
                        var document = new BsonDocument();

                        for (int i = 0; i < headers.Count; i++)
                        {
                            string header = headers[i];
                            string cellValue = row.Cell(i + 1).GetValue<string>(); // Excel is 1-based index
                            document.Add(header, cellValue);
                        }

                        documents.Add(document);
                    }

                    // Insert documents into MongoDB
                    if (documents.Count > 0)
                    {
                        await client.GetDatabase("OBAK").GetCollection<BsonDocument>("Users").InsertManyAsync(documents);
                    }
                }
                var result = await client.GetDatabase("OBAK").GetCollection<BsonDocument>("Users").Find(Builders<BsonDocument>.Filter.Empty).CountDocumentsAsync();
                return Ok(result);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return BadRequest(ex.Message);
            }
        }

    
    }
}
