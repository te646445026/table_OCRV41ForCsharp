namespace table_OCRV41ForCsharp;

public interface IOcrParser
{
    OcrResult Parse(string json);
}