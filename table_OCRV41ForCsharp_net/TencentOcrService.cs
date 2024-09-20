using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;

namespace table_OCRV41ForCsharp;

public class TencentOcrService : IOcrService
{
    private readonly string _secretId;
    private readonly string _secretKey;
    private readonly string _service;
    private readonly string _version;
    private readonly string _action;
    private readonly string _region;
    private static readonly HttpClient Client = new HttpClient();

    public TencentOcrService(string secretId, string secretKey)
    {
        _secretId = secretId;
        _secretKey = secretKey;
        _service = "ocr";
        _version = "2018-11-19";
        _action = "RecognizeTableAccurateOCR";
        _region = "ap-guangzhou";
    }

    public async Task<string> RecognizeTableAsync(string imageBase64)
    {
        var body = imageBase64;
        var token = "";
        var result = await DoRequest(_secretId, _secretKey, _service, _version, _action, body, _region, token);
        return result;
    }

    private async Task<string> DoRequest(string secretId, string secretKey, string service, string version, string action, string body, string region, string token)
    {
        var request = BuildRequest(secretId, secretKey, service, version, action, body, region, token);
        var response = await Client.SendAsync(request);
        return await response.Content.ReadAsStringAsync();
    }

    private HttpRequestMessage BuildRequest(string secretId, string secretKey, string service, string version, string action, string body, string region, string token)
    {
        var host = "ocr.tencentcloudapi.com";
        var url = "https://" + host;
        var contentType = "application/json; charset=utf-8";
        var timestamp = ((int)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalSeconds).ToString();
        var auth = GetAuth(secretId, secretKey, host, contentType, timestamp, body);
        var request = new HttpRequestMessage();
        request.Method = HttpMethod.Post;
        request.Headers.Add("Host", host);
        request.Headers.Add("X-TC-Timestamp", timestamp);
        request.Headers.Add("X-TC-Version", version);
        request.Headers.Add("X-TC-Action", action);
        request.Headers.Add("X-TC-Region", region);
        request.Headers.Add("X-TC-Token", token);
        request.Headers.Add("X-TC-RequestClient", "SDK_NET_BAREBONE");
        request.Headers.TryAddWithoutValidation("Authorization", auth);
        request.RequestUri = new Uri(url);
        request.Content = new StringContent(body, MediaTypeWithQualityHeaderValue.Parse(contentType));
        return request;
    }

    private string GetAuth(string secretId, string secretKey, string host, string contentType, string timestamp, string body)
    {
        var canonicalURI = "/";
        var canonicalHeaders = "content-type:" + contentType + "\nhost:" + host + "\n";
        var signedHeaders = "content-type;host";
        var hashedRequestPayload = Sha256Hex(body);
        var canonicalRequest = "POST" + "\n"
                                          + canonicalURI + "\n"
                                          + "\n"
                                          + canonicalHeaders + "\n"
                                          + signedHeaders + "\n"
                                          + hashedRequestPayload;

        var algorithm = "TC3-HMAC-SHA256";
        var date = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc).AddSeconds(int.Parse(timestamp))
                .ToString("yyyy-MM-dd");
        var service = host.Split(".")[0];
        var credentialScope = date + "/" + service + "/" + "tc3_request";
        var hashedCanonicalRequest = Sha256Hex(canonicalRequest);
        var stringToSign = algorithm + "\n"
                                         + timestamp + "\n"
                                         + credentialScope + "\n"
                                         + hashedCanonicalRequest;

        var tc3SecretKey = Encoding.UTF8.GetBytes("TC3" + secretKey);
        var secretDate = HmacSha256(tc3SecretKey, Encoding.UTF8.GetBytes(date));
        var secretService = HmacSha256(secretDate, Encoding.UTF8.GetBytes(service));
        var secretSigning = HmacSha256(secretService, Encoding.UTF8.GetBytes("tc3_request"));
        var signatureBytes = HmacSha256(secretSigning, Encoding.UTF8.GetBytes(stringToSign));
        var signature = BitConverter.ToString(signatureBytes).Replace("-", "").ToLower();

        return algorithm + " "
                             + "Credential=" + secretId + "/" + credentialScope + ", "
                             + "SignedHeaders=" + signedHeaders + ", "
                             + "Signature=" + signature;
    }

    private static string Sha256Hex(string s)
    {
        using (SHA256 algo = SHA256.Create())
        {
            byte[] hashbytes = algo.ComputeHash(Encoding.UTF8.GetBytes(s));
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < hashbytes.Length; ++i)
            {
                builder.Append(hashbytes[i].ToString("x2"));
            }

            return builder.ToString();
        }
    }

    private static byte[] HmacSha256(byte[] key, byte[] msg)
    {
        using (HMACSHA256 mac = new HMACSHA256(key))
        {
            return mac.ComputeHash(msg);
        }
    }
}