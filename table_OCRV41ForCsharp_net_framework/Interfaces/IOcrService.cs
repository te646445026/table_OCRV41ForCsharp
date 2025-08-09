using table_OCRV41ForCsharp_net_framework.Models;

namespace table_OCRV41ForCsharp_net_framework.Interfaces
{
    public interface IOcrService
    {
        string RecognizeTable(string base64Image, KEY key);
    }
}