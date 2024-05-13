using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Newtonsoft.Json; //https://www.nuget.org/packages/Newtonsoft.Json
using Newtonsoft.Json.Linq;
using System.Collections;
using NPOI.XWPF.UserModel;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using NPOI.SS.Formula.Functions;



namespace table_OCRV41ForCsharp
{
    internal class Program
    {
        // KEY信息
        const string API_KEY = "AKIDCLfBaq2DQVUVbsHoHan5Ml9Slxb5MUVn";
        const string SECRET_KEY = "f9gr9MRp9JIKRRDqMwdSBl9ORZijirto";
        private static readonly HttpClient Client = new HttpClient();

        public class PathMessage
        {
            public string? FolderPath { get; set; }
            public string? DefaultJsonFilePath { get; set; }
            public string? DataFilePath { get; set; }
            public string? DataJsonFilePath { get; set; }
        }

        [STAThread]
        static void Main(string[] args)
        {
            var secretId = API_KEY;
            var secretKey = SECRET_KEY;
            var token = "";
            var service = "ocr";
            var version = "2018-11-19";
            var action = "RecognizeTableAccurateOCR";
            var body = "{}";
            var region = "ap-guangzhou";

            string? workPath ;
            string data_dir = "" ;
            string folder_dir = "";
            
            ArrayList result_dir = new ArrayList();
            Dictionary<string, string> jsonMessage = new Dictionary<string, string>();

            PathMessage path ;

            path = CheckDefaultPath();
            workPath = path.FolderPath;
            

            Console.WriteLine("已生成识别结果请按0，未生成请按1：");
            string? situation = Console.ReadLine();

            if (situation == "1")
            {
                //从json文件中读取
             
                data_dir =  path.DataFilePath+"\\";
                
                folder_dir =  path.DataJsonFilePath+"\\";
               

            }
            else
            {
                // 创建 OpenFileDialog 对象
                OpenFileDialog fileDialog = new OpenFileDialog();
              
                // 设置对话框的属性
                fileDialog.Multiselect = true; // 允许多选文件
                fileDialog.Title = "请选择文件"; // 设置对话框的标题
                fileDialog.Filter = "json文件(*.json)|*.json"; // 设置对话框的文件过滤器

                // 显示对话框并获取用户选择的文件路径
                DialogResult result = fileDialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    foreach(string fileName in fileDialog.FileNames)
                    {
                        result_dir.Add(fileName); // 获取用户选择的多个文件名的数组                                                              // 处理用户选择的文件路径
                    }
                }
            }
            if (situation == "1")
            {
                int num = 0;
                DirectoryInfo directoryInfo = new DirectoryInfo(data_dir);
                foreach(FileInfo file in directoryInfo.GetFiles())
                {
                    Console.WriteLine("{0}: {1} 正在处理：", num + 1, file.Name.Split('.')[0]);
                    body = GetFileContentAsBase64(file.FullName);
                    string data_json = DoRequest(secretId, secretKey, service, version, action, body, region, token);
                    string jsonFile_name = folder_dir + file.Name.Split('.')[0] + ".json";
                    File.WriteAllText(jsonFile_name, data_json);

                    Console.WriteLine("{0}: {1} 下载完成。", num + 1, jsonFile_name);
                    num++;
                    Console.WriteLine("--------------------------------------");
                    Console.WriteLine("");
                    Thread.Sleep(1000);
                }
             
            }
            if(situation == "1")
            {
                string[] file_dir = Directory.GetFiles(folder_dir);
                for (int i = 0; i < file_dir.Length; i++)
                {
                    result_dir.Add(file_dir[i]);
                }

            }
            int fileNum = 0;
            foreach (string jsonPath in result_dir)
            {
                fileNum++;
                Console.WriteLine("-----------{0}-------------", fileNum);
                // 把识别结果的json文档信息提取出来
                jsonMessage = JsonMessage(jsonPath);
                //根据模板，写入对应的word文档里面
                //FileStream docFlieRec = new FileStream(workPath+"\\限速器测试记录模板2.docx",FileMode.OpenOrCreate,FileAccess.ReadWrite);
                //FileStream docFlieRep = new FileStream(workPath+"\\限速器测试报告模板2.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                FileStream docFlieRec = new FileStream(workPath + "\\限速器测试记录模板3.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                FileStream docFlieRep = new FileStream(workPath + "\\限速器测试报告模板3.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite);

                XWPFDocument documentRec = new XWPFDocument(docFlieRec);
                XWPFDocument documentRep = new XWPFDocument(docFlieRep);

                IList<XWPFParagraph> paragraphsRec = documentRec.Paragraphs;
                Console.WriteLine(paragraphsRec[0].ParagraphText + jsonMessage["reportNum2"].ToString());

                
                IList<XWPFTable> tablesRec = documentRec.Tables;
                XWPFTable tableRec0 = tablesRec[0];
                XWPFTable tableRec1 = tablesRec[1];

                IList<XWPFParagraph> paragraphsRep = documentRep.Paragraphs;
                Console.WriteLine(paragraphsRep[0].ParagraphText + jsonMessage["reportNum2"].ToString());

                IList<XWPFTable> tablesRep = documentRep.Tables;
                XWPFTable tableRep0 = tablesRep[0];
                XWPFTable tableRep1 = tablesRep[1];


                
                //写入记录for模板3
                try
                {
                    tableRec1.GetRow(0).GetCell(1).SetText(jsonMessage["userName"]);
                    //左对齐
                    tableRec1.GetRow(0).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;

                }
                catch
                {
                    Console.WriteLine("userName write error");
                }

                try
                {
                    tableRec1.GetRow(1).GetCell(1).SetText(jsonMessage["userName"]);
                    //左对齐
                    tableRec1.GetRow(1).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("userName write error");
                }

                try
                {
                    tableRec1.GetRow(2).GetCell(1).SetText(jsonMessage["deviceCode"]);
                    //左对齐
                    tableRec1.GetRow(2).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("deviceCode write error");
                }

                try
                {
                    tableRec1.GetRow(3).GetCell(1).SetText(jsonMessage["serialNum"]);
                    //左对齐
                    tableRec1.GetRow(3).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("serialNum write error");
                }

                try
                {
                    tableRec1.GetRow(4).GetCell(2).SetText(jsonMessage["xiansuqiModel"]);
                    //左对齐
                    tableRec1.GetRow(4).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("xiansuqiModel write error");
                }

                try
                {
                    tableRec1.GetRow(4).GetCell(4).SetText(jsonMessage["xiansuqiNum"]);
                    //左对齐
                    tableRec1.GetRow(4).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("xiansuqiNum write error");
                }

                try
                {
                    tableRec1.GetRow(5).GetCell(2).SetText(jsonMessage["speed"] + "m/s");
                    //左对齐
                    tableRec1.GetRow(5).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("speed write error");
                }

                try
                {
                    tableRec1.GetRow(5).GetCell(4).SetText(jsonMessage["xiansuqiDirection"]);
                    //左对齐
                    tableRec1.GetRow(5).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("direction write error");
                }

                try
                {
                    tableRec1.GetRow(15).GetCell(1).SetText(jsonMessage["temperature"]);
                    //左对齐
                    tableRec1.GetRow(15).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("direction write error");
                }

                try
                {
                    tableRec1.GetRow(17).GetCell(1).SetText(jsonMessage["MaintenanceUnit"]);
                    //左对齐
                    tableRec1.GetRow(17).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("MaintenanceUnit write error");
                }

                try
                {
                    tableRec1.GetRow(19).GetCell(3).SetText(jsonMessage["next_year"]);
                    //右对齐
                    tableRec1.GetRow(19).GetCell(3).Paragraphs[0].Alignment = ParagraphAlignment.RIGHT;
                }
                catch
                {
                    Console.WriteLine("nextYear write error");
                }

                try
                {
                    tableRec1.GetRow(18).GetCell(3).Paragraphs[0].CreateRun().SetText(jsonMessage["date"]);
                    //右对齐
                    tableRec1.GetRow(18).GetCell(3).Paragraphs[0].Alignment = ParagraphAlignment.RIGHT;

                }
                catch
                {
                    Console.WriteLine("date write error");
                }

                try
                {
                    if (jsonMessage["jianyanOrjiance"].Equals("检验"))
                    {
                        paragraphsRec[0].CreateRun().SetText("D");
                    }
                    else
                    {
                        paragraphsRec[0].CreateRun().SetText("E");
                    }
                    paragraphsRec[0].CreateRun().SetText(jsonMessage["reportNum2"]);
                    
                }
                catch
                {
                    Console.WriteLine("reportNum2 write error");
                }


                string outPath = string.Format(workPath+"\\{0}_{1}_{2}.doc",
                                                    jsonMessage["deviceCode"], 
                                                    Path.GetFileNameWithoutExtension(jsonPath),
                                                    jsonMessage["next_year_flag"]);
                FileStream outFile = new FileStream(outPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                documentRec.Write(outFile);
                outFile.Close();
                documentRec.Close();
                docFlieRec.Close();

                Console.WriteLine("{0}打印记录完成", Path.GetFileNameWithoutExtension(jsonPath));
                Console.WriteLine("-------------------------------------------------------");
                Console.WriteLine("");
                
                //写入报告模板3
                try
                {
                    tableRep1.GetRow(0).GetCell(1).SetText(jsonMessage["userName"]);
                    //左对齐
                    tableRep1.GetRow(0).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("userName write error");
                }

                try
                {
                    tableRep1.GetRow(1).GetCell(1).SetText(jsonMessage["userName"]);
                    //左对齐
                    tableRep1.GetRow(1).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("userName write error");
                }

                try
                {
                    tableRep1.GetRow(2).GetCell(1).SetText(jsonMessage["deviceCode"]);
                    //左对齐
                    tableRep1.GetRow(2).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("deviceCode write error");
                }

                try
                {
                    tableRep1.GetRow(3).GetCell(1).SetText(jsonMessage["serialNum"]);
                    //左对齐
                    tableRep1.GetRow(3).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("serialNum write error");
                }

                try
                {
                    tableRep1.GetRow(4).GetCell(2).SetText(jsonMessage["xiansuqiModel"]);
                    //左对齐
                    tableRep1.GetRow(4).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("xiansuqiModel write error");
                }

                try
                {
                    tableRep1.GetRow(4).GetCell(4).SetText(jsonMessage["xiansuqiNum"]);
                    //左对齐
                    tableRep1.GetRow(4).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("xiansuqiNum write error");
                }

                try
                {
                    tableRep1.GetRow(5).GetCell(2).SetText(jsonMessage["speed"] + "m/s");
                    //左对齐
                    tableRep1.GetRow(5).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("speed write error");
                }

                try
                {
                    tableRep1.GetRow(5).GetCell(4).SetText(jsonMessage["xiansuqiDirectionForReport"]);
                    //左对齐
                    tableRep1.GetRow(5).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("direction write error");
                }

                

                try
                {
                    tableRep1.GetRow(15).GetCell(1).SetText(jsonMessage["MaintenanceUnit"]);
                    //右对齐
                    tableRep1.GetRow(15).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                }
                catch
                {
                    Console.WriteLine("MaintenanceUnit write error");
                }

                try
                {
                    tableRep1.GetRow(16).GetCell(3).SetText(jsonMessage["next_year"]);
                    //右对齐
                    tableRep1.GetRow(16).GetCell(3).Paragraphs[0].Alignment = ParagraphAlignment.RIGHT;
                }
                catch
                {
                    Console.WriteLine("nextYear write error");
                }

                try
                {
                    tableRep1.GetRow(17).GetCell(0).Paragraphs[0].CreateRun().SetText(jsonMessage["date"]);


                    tableRep1.GetRow(18).GetCell(0).Paragraphs[0].CreateRun().SetText(jsonMessage["shenhe_date"]);
                    tableRep1.GetRow(19).GetCell(0).Paragraphs[0].CreateRun().SetText(jsonMessage["shenhe_date"]);


                    //左对齐
                    tableRep1.GetRow(17).GetCell(0).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                    tableRep1.GetRow(18).GetCell(0).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
                    tableRep1.GetRow(19).GetCell(0).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;

                }
                catch
                {
                    Console.WriteLine("date write error");
                }

                try
                {
                    if (jsonMessage["jianyanOrjiance"].Equals("检验"))
                    {
                        paragraphsRep[0].CreateRun().SetText("D");
                    }
                    else
                    {
                        paragraphsRep[0].CreateRun().SetText("E");
                    }
                    paragraphsRep[0].CreateRun().SetText(jsonMessage["reportNum2"]);
                    //if (jsonMessage["xiansuqiDirectionForReport"] == "双向")
                    //{
                    //    paragraphsRep[0].CreateRun().SetText("D");
                    //}
                }
                catch
                {
                    Console.WriteLine("reportNum2 write error");
                }


                string outPath2 = string.Format(workPath + "\\{0}.doc",
                                                    jsonMessage["deviceCode"]);
                FileStream outFile2 = new FileStream(outPath2, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                documentRep.Write(outFile2);
                outFile2.Close();
                documentRep.Close();
                docFlieRep.Close();

                Console.WriteLine("{0}打印报告完成", Path.GetFileNameWithoutExtension(jsonPath));
                Console.WriteLine("-------------------------------------------------------");
                Console.WriteLine("");

            }
            Console.WriteLine("输入任意按钮退出");
            Console.ReadKey();
        }

        /**
        * 获取文件base64编码
        * @param path 文件路径
        * @return base64编码信息，不带文件头
        */
        static string GetFileContentAsBase64(string path)
        {
            using (FileStream filestream = new FileStream(path, FileMode.Open))
            {
                byte[] arr = new byte[filestream.Length];
                filestream.Read(arr, 0, (int)filestream.Length);
                string base64 = Convert.ToBase64String(arr);
                base64 = "{\"ImageBase64\":\"data:image/png;base64," + base64 + "\"}";
                return base64;
            }
        }

        static string DoRequest(
            string secretId, string secretKey,
            string service, string version, string action,
            string body, string region, string token)
        {
            var request = BuildRequest(secretId, secretKey, service, version, action, body, region, token);
            var response = Client.Send(request);
            return response.Content.ReadAsStringAsync().Result;
        }

        static HttpRequestMessage BuildRequest(
            string secretId, string secretKey,
            string service, string version, string action,
            string body, string region, string token)
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
            // request.Headers.Authorization = new AuthenticationHeaderValue(auth);
            request.RequestUri = new Uri(url);
            request.Content = new StringContent(body, MediaTypeWithQualityHeaderValue.Parse(contentType));
            //Console.WriteLine(request);
            return request;
        }

        static string GetAuth(
            string secretId, string secretKey, 
            string host, string contentType,
            string timestamp, string body)
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

        public static string Sha256Hex(string s)
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

        /**
       * 检查默认配置文件，如果存在，则读取对应部分作为默认路径，如果不存在，则新建并要求用户选择路径作为默认路径
       * @param 可选参数defaultPath 默认为调试用的路径
       * @return 返回一个对象，FolderPath为工作路径，DefaultJsonFilePath为默认参数的json文件路径，
       *                DataFilePath为截图保存文件的路径,DataJsonFilePath为识别结果保存的文件路径
       */
        static PathMessage CheckDefaultPath(string defaultFolderPath = "")
        {
            PathMessage path = new PathMessage();
                        
           
            //检查是否有配置文件，没有就生成，并选择路径
            path.DefaultJsonFilePath = System.Environment.CurrentDirectory + @"\default.json";
         

            if (!File.Exists(path.DefaultJsonFilePath))
            {               
                //选择默认路径
                FolderBrowserDialog folder1 = new FolderBrowserDialog();
                folder1.Description = "请选择需要识别的图片所在文件夹";

                if (folder1.ShowDialog() == DialogResult.OK)
                {
                    path.DataFilePath = folder1.SelectedPath ;
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine($"已选择需要识别的图片所在文件夹：{path.DataFilePath}");
                    Console.WriteLine("");
                }

                FolderBrowserDialog folder2 = new FolderBrowserDialog();
                folder2.Description = "请选择识别结果存放的文件夹";

                if (folder2.ShowDialog() == DialogResult.OK)
                {
                    path.DataJsonFilePath = folder2.SelectedPath ;
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine($"已选择识别结果存放的文件夹：{path.DataJsonFilePath}");
                    Console.WriteLine("");
                }
                //确定工作路径

                path.FolderPath = Path.GetDirectoryName(folder1 .SelectedPath);
 
                Console.WriteLine($"当前工作文件夹目录为：{path.FolderPath}");

                //把默认路径写入json文件中
                string defaultStr = JsonConvert.SerializeObject(path);
                File.WriteAllText(path.DefaultJsonFilePath, defaultStr);                            
            }
            else
            {
                //读取DataFilePath和DataJsonFilePath
                string defaultStr = File.ReadAllText(path.DefaultJsonFilePath);
                PathMessage? defaultStrPath = JsonConvert.DeserializeObject<PathMessage>(defaultStr);
                path.DataFilePath = defaultStrPath.DataFilePath;
                path.DataJsonFilePath= defaultStrPath.DataJsonFilePath;
                //确定工作路径
                path.FolderPath = Path.GetDirectoryName(path.DataFilePath);
                Console.WriteLine($"当前工作文件夹目录为：{path.FolderPath}");
            }

            return  path;
        }

        static void ObjsIndex(string str,JObject objs,out int indexj,out int indexi,out bool isContain)
        {
            
            indexi = 0;
            indexj = 0;
            isContain = false;

            for (int j = 0; j < objs["Response"]["TableDetections"].Count(); j++)
            {
                for (int i = 0; i < objs["Response"]["TableDetections"][j]["Cells"].Count(); i++)
                {
                    var text = objs["Response"]["TableDetections"][j]["Cells"][i]["Text"];
                    isContain = text.ToString().Contains(str);
                    if (isContain)
                    {
                        indexi = i;
                        indexj = j;
                        return;
                    }
                }          
            }
        }

        static Dictionary<string, string> JsonMessage(string filePath)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            string json = File.ReadAllText(filePath);
            JObject? objs = JObject.Parse(json);


            string jianyanOrjiance = "检测";
            try
            {
                
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("RTD", objs, out indexj, out indexi, out isContain);

                if (isContain)
                {
                    jianyanOrjiance = "检验";                   
                }


                Console.WriteLine("当前图片是: " + jianyanOrjiance);
            }
            catch
            {
                Console.WriteLine("获取检验还是检测失败,默认设置为检测" );

            }
            
            
            
            string deviceCode;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("设备代码", objs, out indexj, out indexi, out isContain);

                deviceCode = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                
                Console.WriteLine("设备代码: " + deviceCode);
            }
            catch
            {
                Console.WriteLine("设备代码获取错误");
                deviceCode = "/";
            }
            string model;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("型号", objs, out indexj, out indexi, out isContain);
                model = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("型号: " + model);
            }
            catch
            {
                Console.WriteLine("型号获取错误");
                model = "/";
            }
            string serialNum;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("产品编号", objs, out indexj, out indexi, out isContain);
                serialNum = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("产品编号: " + serialNum);
            }
            catch
            {
                Console.WriteLine("产品编号获取错误");
                serialNum = "/";
            }
            string ManufacturingUnit;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("制造单位", objs, out indexj, out indexi, out isContain);
                ManufacturingUnit = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("制造单位: " + ManufacturingUnit);
            }
            catch
            {
                Console.WriteLine("制造单位获取错误");
                ManufacturingUnit = "/";
            }
            string userName;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("使用单位", objs, out indexj, out indexi, out isContain);
                userName = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("使用单位: " + userName);
            }
            catch
            {
                Console.WriteLine("使用单位获取错误");
                userName = "/";
            }
            string UsingAddress;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("安装地点", objs, out indexj, out indexi, out isContain);
                UsingAddress = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("安装地点: " + UsingAddress);
            }
            catch
            {
                Console.WriteLine("安装地点获取错误");
                UsingAddress = "/";
            }
            string MaintenanceUnit;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("维护保养单位", objs, out indexj, out indexi, out isContain);
                MaintenanceUnit = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("维护保养单位: " + MaintenanceUnit);
            }
            catch
            {
                Console.WriteLine("维护保养单位获取错误");
                MaintenanceUnit = "/";
            }
            string speed;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("额定速度", objs, out indexj, out indexi, out isContain);
                speed = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "");
                string speed_pattern = @"\d{1}.\d{1,2}";
                var speedNeed = Regex.Matches(speed, speed_pattern);
                speed = speedNeed[0].ToString();
                Console.WriteLine("速度：" + speed);
            }
            catch
            {
                Console.WriteLine("速度获取错误");
                speed = "/";
            }
            string temperature;
            try
            {
                
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("条件", objs, out indexj, out indexi, out isContain);
                temperature = objs["Response"]["TableDetections"][indexj]["Cells"][indexi+1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                string temperature_pattern = @"\d{2,3}";
                MatchCollection temperatureNeed = Regex.Matches(temperature, temperature_pattern);
                temperature = $"温度：{temperatureNeed[0].ToString()}℃，  湿度：{temperatureNeed[1].ToString()}％ ， 电压：{temperatureNeed[2].ToString()}V";
                Console.WriteLine("温度、湿度、电压: " + temperature);
            }
            catch
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("条件", objs, out indexj, out indexi, out isContain);
                temperature = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                string temperature2 = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 3]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                string temperature3 = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 5]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                temperature = $"温度：{temperature}℃，  湿度：{temperature2}％ ， 电压：{temperature3}V";
                Console.WriteLine("温度、湿度、电压: " + temperature);
            }

            string reportNum;
            string reportNum2;
            string jianyanOrjianceReportNum;
            try
            {
                if (jianyanOrjiance.Equals("检验"))
                {
                    jianyanOrjianceReportNum = "RTD";
                }
                else
                {
                    jianyanOrjianceReportNum = "RTE";
                }
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex(jianyanOrjianceReportNum, objs, out indexj, out indexi, out isContain);

                reportNum = objs["Response"]["TableDetections"][indexj]["Cells"][indexi]["Text"].ToString();
                //MatchCollection matchs = Regex.Matches(reportNum, @"^\d{8}");
                //reportNum2 = matchs[0].ToString().Substring(1,7);
                reportNum2 = reportNum.Substring(reportNum.Length - 7);
                Console.WriteLine("报告编号: " + reportNum2);

            }
            catch
            {
                Console.WriteLine("报告编号获取错误");
                reportNum = "/";
                reportNum2 = "/";
            }
            string? date;
            string next_year;
            string next_year_flag;
            string shenhe_date;
            string jianyanOrjianceDate;
            try
            {
                if (jianyanOrjiance.Equals("检验"))
                {
                    jianyanOrjianceDate = "检验日期";
                }
                else
                {
                    jianyanOrjianceDate = "检测日期";
                }

                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex(jianyanOrjianceDate, objs, out indexj, out indexi, out isContain);
                date = objs["Response"]["TableDetections"][indexj]["Cells"][indexi]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                string date_or_month_pattern = @"\d{4}年\d{1,2}[\u4e00-\u9fa5]\d{0,}日|\d{4}年\d{1,2}[\u4e00-\u9fa5]";
                MatchCollection dateNeed = Regex.Matches(date, date_or_month_pattern);
                if (dateNeed != null)
                {
                    date = dateNeed[dateNeed.Count()-1].Value;
                    string date_or_month_pattern2 = @"\d+";
                    MatchCollection matches = Regex.Matches(date, date_or_month_pattern2);
                    int year = int.Parse(matches[0].Value);
                    int month = int.Parse(matches[1].Value);
                    int day = int.Parse(matches[2].Value);
                    date = matches[0].Value + "年" + matches[1].Value + "月" + matches[2].Value + "日";
                    Console.WriteLine("检验时间为："+date);
                    DateTime dateforcell = new DateTime(year, month, day);
                    //计算2年后的日期
                    string? nextdate;
                    Console.WriteLine("请输入下次检验日期间隔，1代表1年，2代表2年: ");
                    nextdate = Console.ReadLine();
                    while (nextdate != "1" & nextdate != "2")
                    {
                        Console.WriteLine("请输入下次检验日期间隔，1代表1年，2代表2年: ");
                        nextdate = Console.ReadLine();
                    }

                    DateTime next_year_date = dateforcell.AddYears(int.Parse(nextdate));
                    next_year = next_year_date.ToString("yyyy年MM月dd日");
                    next_year_flag = "";
                    //计算审核校准日期
                    DateTime shenhe_dateforcell = dateforcell.AddDays(1);
                    shenhe_date = shenhe_dateforcell.ToString("yyyy年MM月dd日");
                }
                else
                {
                    date = "   年   月   日";
                    next_year = "   年   月   日";
                    next_year_flag = "检验日期和下检日期出错";
                    shenhe_date = "   年   月   日";
                    Console.WriteLine("检验日期获取错误");
                }               
            }
            catch
            {
                Console.WriteLine("检验日期获取错误");
                date = "   年   月   日";
                next_year = "   年   月   日";
                next_year_flag = "检验日期和下检日期出错";
                shenhe_date = "   年   月   日";
            }
                                 
            string xiansuqiModel;
            try
            {
                Console.WriteLine("输入限速器型号：");
                xiansuqiModel = Console.ReadLine();
                Console.WriteLine("限速器型号：" + xiansuqiModel);
            }
            catch
            {
                Console.WriteLine("限速器型号获取错误");
                xiansuqiModel = "/";
            }
            string xiansuqiNum;
            try
            {
                Console.WriteLine("输入限速器编号：");
                xiansuqiNum = Console.ReadLine();
                Console.WriteLine("限速器编号：" + xiansuqiNum);
            }
            catch
            {
                Console.WriteLine("限速器编号获取错误");
                xiansuqiNum = "/";
            }
            string xiansuqiDirection;
            string xiansuqiDirectionForReport;
            Console.WriteLine("输入单向还是双向，0为单向，1为双向");
            if (Console.ReadLine() == "0")
            {
                xiansuqiDirection = "☑  单向 ☐  双向";
                xiansuqiDirectionForReport = "单向";
            }
            else
            {
                xiansuqiDirection = "☐  单向 ☑  双向";
                xiansuqiDirectionForReport = "双向";
            }

            result.Add("userName", userName);
            result.Add("MaintenanceUnit", MaintenanceUnit);
            result.Add("ManufacturingUnit", ManufacturingUnit);
            result.Add("UsingAddress", UsingAddress);
            result.Add("deviceCode", deviceCode);
            result.Add("model", model);
            result.Add("serialNum", serialNum);
            result.Add("speed", speed);
            result.Add("xiansuqiModel", xiansuqiModel);
            result.Add("xiansuqiNum", xiansuqiNum);
            result.Add("reportNum2", reportNum2);
            result.Add("date", date);
            result.Add("next_year", next_year);
            result.Add("xiansuqiDirection", xiansuqiDirection);
            result.Add("xiansuqiDirectionForReport", xiansuqiDirectionForReport);
            result.Add("next_year_flag", next_year_flag);
            result.Add("shenhe_date", shenhe_date);
            result.Add("temperature", temperature);
            result.Add("jianyanOrjiance", jianyanOrjiance);

            return result;
        }
    }
}
