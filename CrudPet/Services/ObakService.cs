using MongoDB.Bson;
using MongoDB.Driver;
using ClosedXML.Excel;
using System.IO;
using System.Threading.Tasks;

public class ObakService : IObakService
{
    private readonly IMongoClient _client;
    private readonly IMongoDatabase _database;

    public ObakService(string connectionString)
    {
        var settings = MongoClientSettings.FromConnectionString(connectionString);
        settings.ServerApi = new ServerApi(ServerApiVersion.V1);
        _client = new MongoClient(settings);
        _database = _client.GetDatabase("OBAK");
    }

    public async Task<long> GetAndInsertDataAsync(string filePath)
    {
        try
        {
            using (Stream xlsxStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var workbook = new XLWorkbook(xlsxStream);
                var worksheet = workbook.Worksheet(1);

                var headerRow = worksheet.Row(1);
                var headers = new List<string>();

                foreach (var headerCell in headerRow.CellsUsed())
                {
                    headers.Add(headerCell.Value.ToString());
                }

                var documents = new List<BsonDocument>();

                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    var document = new BsonDocument();

                    for (int i = 0; i < headers.Count; i++)
                    {
                        string header = headers[i];
                        string cellValue = row.Cell(i + 1).GetValue<string>();
                        document.Add(header, cellValue);
                    }

                    documents.Add(document);
                }

                if (documents.Count > 0)
                {
                    await _database.GetCollection<BsonDocument>("Users").InsertManyAsync(documents);
                }
            }

            return await _database.GetCollection<BsonDocument>("Users").CountDocumentsAsync(Builders<BsonDocument>.Filter.Empty);
        }
        catch (Exception ex)
        {
            throw new Exception("An error occurred while processing the data.", ex);
        }
    }
}
