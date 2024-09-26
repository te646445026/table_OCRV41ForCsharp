using Newtonsoft.Json.Linq;

namespace table_OCRV41ForCsharp;

public class TencentOcrParser:IOcrParser
{
    public OcrResult Parse(string json)
    {
        var objs = JObject.Parse(json);
        OcrResult result = new OcrResult();



    }
}