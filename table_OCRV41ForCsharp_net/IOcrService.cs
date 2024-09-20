namespace table_OCRV41ForCsharp;

public interface IOcrService
{
    Task<string> RecognizeTableAsync(string imageBase64);
}