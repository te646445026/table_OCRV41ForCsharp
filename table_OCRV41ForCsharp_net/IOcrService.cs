namespace table_OCRV41ForCsharp_net;

public interface IOcrService
{
    Task<string> RecognizeTableAsync(string imageBase64);
}