using table_OCRV41ForCsharp_net_framework.Models;

namespace table_OCRV41ForCsharp_net_framework.Interfaces
{
    public interface IOcrParser
    {
        OcrResult Parse(string ocrResponse);
    }
}