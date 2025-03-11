public interface IObakService
{
    Task<long> GetAndInsertDataAsync(string filePath);
}
