using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using RestSharp;//依赖版本106.15.0 https://www.nuget.org/packages/RestSharp/106.15.0
using Newtonsoft.Json; //https://www.nuget.org/packages/Newtonsoft.Json
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Collections;
using NPOI.XWPF.UserModel;
using Org.BouncyCastle.Asn1.X509;

namespace table_OCRV41ForCsharp
{
    internal class Program
    {
        // KEY信息
        const string API_KEY = "Et4nGdx8ecc5chOnoilbxEyX";
        const string SECRET_KEY = "505cd0eiUZt22mPzelDGVrWzN7ELwteh";
        const string REQUEST_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/table";
        
        [STAThread]
        static void Main(string[] args)
        {
            string workPath = "";
            Console.WriteLine("已生成识别结果请按0，未生成请按1：");
            string situation = Console.ReadLine();

            string data_dir="" ;
            string folder_dir = "";
            ArrayList result_dir = new ArrayList();
            Dictionary<string, string> jsonMessage = new Dictionary<string, string>();

            if (situation == "1")
            {
                FolderBrowserDialog folder1 = new FolderBrowserDialog();
                folder1.Description = "请选择需要识别的图片所在文件夹";

                if (folder1.ShowDialog() == DialogResult.OK)
                {
                    data_dir = folder1.SelectedPath + "\\";
                }
                FolderBrowserDialog folder2 = new FolderBrowserDialog();
                folder2.Description = "请选择识别结果存放的文件夹";

                if (folder2.ShowDialog() == DialogResult.OK)
                {
                    folder_dir = folder2.SelectedPath + "\\";
                }

                workPath = Path.GetDirectoryName(folder1.SelectedPath);
                Console.WriteLine("当前工作路径为：" + workPath);

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

                workPath = Path.GetDirectoryName( Path.GetDirectoryName(fileDialog.FileNames[0]));
                Console.WriteLine("当前工作路径为：" + workPath);
            }
            if (situation == "1")
            {
                int num = 0;
                DirectoryInfo directoryInfo = new DirectoryInfo(data_dir);
                foreach(FileInfo file in directoryInfo.GetFiles())
                {
                    Console.WriteLine("{0}: {1} 正在处理：", num + 1, file.Name.Split('.')[0]);
                    string data_json = BaiduApi(file.FullName, REQUEST_URL, GetAccessToken());
                    string jsonFile_name = folder_dir + file.Name.Split('.')[0] + ".json";
                    File.WriteAllText(jsonFile_name, data_json);

                    Console.WriteLine("{0}: {1} 下载完成。", num + 1, jsonFile_name);
                    num++;
                    Console.WriteLine("--------------------------------------");
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
                FileStream docFlieRec = new FileStream(workPath+"\\限速器测试记录模板2.docx",FileMode.OpenOrCreate,FileAccess.ReadWrite);
                FileStream docFlieRep = new FileStream(workPath+"\\限速器测试报告模板2.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite);

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


                //写入记录
                try
                {
                    tableRec1.GetRow(0).GetCell(1).SetText(jsonMessage["userName"]);
                }
                catch
                {
                    Console.WriteLine("userName write error");
                }

                try
                {
                    tableRec1.GetRow(1).GetCell(1).SetText(jsonMessage["MaintenanceUnit"]);
                }
                catch
                {
                    Console.WriteLine("MaintenanceUnit write error");
                }

                try
                {
                    tableRec1.GetRow(2).GetCell(1).SetText(jsonMessage["ManufacturingUnit"]);
                }
                catch
                {
                    Console.WriteLine("ManufacturingUnit write error");
                }

                try
                {
                    tableRec1.GetRow(3).GetCell(1).SetText(jsonMessage["UsingAddress"]);
                }
                catch
                {
                    Console.WriteLine("UsingAddress write error");
                }

                try
                {
                    tableRec1.GetRow(4).GetCell(1).SetText(jsonMessage["deviceCode"]);
                }
                catch
                {
                    Console.WriteLine("deviceCode write error");
                }

                try
                {
                    tableRec1.GetRow(5).GetCell(1).SetText(jsonMessage["model"]);
                }
                catch
                {
                    Console.WriteLine("model write error");
                }

                try
                {
                    tableRec1.GetRow(5).GetCell(3).SetText(jsonMessage["serialNum"]);
                }
                catch
                {
                    Console.WriteLine("serialNum write error");
                }

                try
                {
                    tableRec1.GetRow(6).GetCell(2).SetText(jsonMessage["xiansuqiModel"]);
                }
                catch
                {
                    Console.WriteLine("xiansuqiModel write error");
                }

                try
                {
                    tableRec1.GetRow(6).GetCell(4).SetText(jsonMessage["xiansuqiNum"]);
                }
                catch
                {
                    Console.WriteLine("xiansuqiNum write error");
                }

                try
                {
                    tableRec1.GetRow(7).GetCell(2).SetText(jsonMessage["speed"] + "m/s");
                }
                catch
                {
                    Console.WriteLine("speed write error");
                }

                try
                {
                    tableRec1.GetRow(7).GetCell(4).SetText(jsonMessage["xiansuqiDirectionForReport"]);
                }
                catch
                {
                    Console.WriteLine("direction write error");
                }

                try
                {
                    tableRec1.GetRow(21).GetCell(3).SetText(jsonMessage["next_year"]);
                }
                catch
                {
                    Console.WriteLine("nextYear write error");
                }

                try
                {
                    tableRec1.GetRow(22).GetCell(1).Paragraphs[0].CreateRun().SetText(jsonMessage["date"]);
                    tableRec1.GetRow(22).GetCell(3).Paragraphs[0].CreateRun().SetText(jsonMessage["date"]);
                }
                catch
                {
                    Console.WriteLine("date write error");
                }

                try
                {
                    paragraphsRec[0].CreateRun().SetText(jsonMessage["reportNum2"]);
                    if (jsonMessage["xiansuqiDirectionForReport"] == "双向")
                    {
                        paragraphsRec[0].CreateRun().SetText("D");
                    }
                }
                catch
                {
                    Console.WriteLine("reportNum2 write error");
                }


                string outPath = string.Format(workPath+"\\{0}_{1}_{2}.docx",
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
                //写入报告
                try
                {
                    tableRep1.GetRow(0).GetCell(1).SetText(jsonMessage["userName"]);
                }
                catch
                {
                    Console.WriteLine("userName write error");
                }

                try
                {
                    tableRep1.GetRow(1).GetCell(1).SetText(jsonMessage["MaintenanceUnit"]);
                }
                catch
                {
                    Console.WriteLine("MaintenanceUnit write error");
                }

                try
                {
                    tableRep1.GetRow(2).GetCell(1).SetText(jsonMessage["ManufacturingUnit"]);
                }
                catch
                {
                    Console.WriteLine("ManufacturingUnit write error");
                }

                try
                {
                    tableRep1.GetRow(3).GetCell(1).SetText(jsonMessage["UsingAddress"]);
                }
                catch
                {
                    Console.WriteLine("UsingAddress write error");
                }

                try
                {
                    tableRep1.GetRow(4).GetCell(1).SetText(jsonMessage["deviceCode"]);
                }
                catch
                {
                    Console.WriteLine("deviceCode write error");
                }

                try
                {
                    tableRep1.GetRow(5).GetCell(1).SetText(jsonMessage["model"]);
                }
                catch
                {
                    Console.WriteLine("model write error");
                }

                try
                {
                    tableRep1.GetRow(5).GetCell(3).SetText(jsonMessage["serialNum"]);
                }
                catch
                {
                    Console.WriteLine("serialNum write error");
                }

                try
                {
                    tableRep1.GetRow(6).GetCell(2).SetText(jsonMessage["xiansuqiModel"]);
                }
                catch
                {
                    Console.WriteLine("xiansuqiModel write error");
                }

                try
                {
                    tableRep1.GetRow(6).GetCell(4).SetText(jsonMessage["xiansuqiNum"]);
                }
                catch
                {
                    Console.WriteLine("xiansuqiNum write error");
                }

                try
                {
                    tableRep1.GetRow(7).GetCell(2).SetText(jsonMessage["speed"] + "m/s");
                }
                catch
                {
                    Console.WriteLine("speed write error");
                }

                try
                {
                    tableRep1.GetRow(7).GetCell(4).SetText(jsonMessage["xiansuqiDirectionForReport"]);
                }
                catch
                {
                    Console.WriteLine("direction write error");
                }

                try
                {
                    tableRep1.GetRow(21).GetCell(3).SetText(jsonMessage["next_year"]);
                }
                catch
                {
                    Console.WriteLine("nextYear write error");
                }

                try
                {
                    tableRep1.GetRow(22).GetCell(1).Paragraphs[0].CreateRun().SetText(jsonMessage["date"]);

                    
                    tableRep1.GetRow(23).GetCell(1).Paragraphs[0].CreateRun().SetText(jsonMessage["shenhe_date"]);
                    tableRep1.GetRow(24).GetCell(1).Paragraphs[0].CreateRun().SetText(jsonMessage["shenhe_date"]);
                    tableRep1.GetRow(22).GetCell(2).Paragraphs[3].CreateRun().SetText(jsonMessage["shenhe_date"]);
                }
                catch
                {
                    Console.WriteLine("date write error");
                }

                try
                {
                    paragraphsRep[0].CreateRun().SetText(jsonMessage["reportNum2"]);
                    if (jsonMessage["xiansuqiDirectionForReport"] == "双向")
                    {
                        paragraphsRep[0].CreateRun().SetText("D");
                    }
                }
                catch
                {
                    Console.WriteLine("reportNum2 write error");
                }


                string outPath2 = string.Format(workPath + "\\{0}.docx",
                                                    jsonMessage["deviceCode"]);
                FileStream outFile2 = new FileStream(outPath2, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                documentRep.Write(outFile2);
                outFile2.Close();
                documentRep.Close();
                docFlieRep.Close();

                Console.WriteLine("{0}打印报告完成", Path.GetFileNameWithoutExtension(jsonPath));
                Console.WriteLine("-------------------------------------------------------");

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
                return base64;
            }
        }

        /**
        * 使用 AK，SK 生成鉴权签名（Access Token）
        * @return 鉴权签名信息（Access Token）
        */
        static string GetAccessToken()
        {
            var client = new RestClient($"https://aip.baidubce.com/oauth/2.0/token");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddParameter("grant_type", "client_credentials");
            request.AddParameter("client_id", API_KEY);
            request.AddParameter("client_secret", SECRET_KEY);
            IRestResponse response = client.Execute(request);
            var result = JsonConvert.DeserializeObject<dynamic>(response.Content);
            //var result = JsonSerializer.Deserialize<dynamic>(response.Content);
            return result.access_token.ToString();
        }

        /**
         * 获取识别的结果，以json文件的形式返回字符串
         * @param filePath 需要识别的图片路径
         * @param requestUrl 上传图片的url
         * @param acessToken  Token文件
         */
        static string BaiduApi(string filePath, string requestUrl, string accessToken)
        {
            // image 可以通过 GetFileBase64Content('C:\fakepath\双龙.png') 方法获取
            string pictureBase64 = GetFileContentAsBase64(filePath);
            //开启 post服务
            var client = new RestClient(requestUrl + "?access_token=" + accessToken);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddHeader("Accept", "application/json");
            request.AddParameter("image", pictureBase64);
            request.AddParameter("cell_contents", "false");
            request.AddParameter("return_excel", "false");
            IRestResponse response = client.Execute(request);
            return response.Content;
        }

        static Dictionary<string, string> JsonMessage(string filePath)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            string json = File.ReadAllText(filePath);
            JObject objs = JObject.Parse(json);
            string deviceCode;
            try
            {
                deviceCode = objs["tables_result"][0]["body"][1]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                model = objs["tables_result"][0]["body"][7]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                serialNum = objs["tables_result"][0]["body"][9]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                ManufacturingUnit = objs["tables_result"][0]["body"][13]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                userName = objs["tables_result"][0]["body"][15]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                UsingAddress = objs["tables_result"][0]["body"][21]["words"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("使用地点: " + UsingAddress);
            }
            catch
            {
                Console.WriteLine("使用地点获取错误");
                UsingAddress = "/";
            }
            string MaintenanceUnit;
            try
            {
                MaintenanceUnit = objs["tables_result"][0]["body"][31]["words"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("维保单位: " + MaintenanceUnit);
            }
            catch
            {
                Console.WriteLine("维保单位获取错误");
                MaintenanceUnit = "/";
            }
            string reportNum;
            string reportNum2;
            try
            {
                reportNum = objs["tables_result"][0]["header"][1]["words"].ToString();
                MatchCollection matchs = Regex.Matches(reportNum, @"\d{7}");
                reportNum2 = matchs[0].ToString();
                Console.WriteLine("报告编号: " + reportNum2);
            }
            catch
            {
                Console.WriteLine("报告编号获取错误");
                reportNum = "/";
                reportNum2 = "/";
            }
            string date;
            string next_year;
            string next_year_flag;
            string shenhe_date;
            try
            {
                date = objs["tables_result"][0]["body"][56]["words"].ToString().Replace("\n", "").Replace("\r", "");
                if (!date.Contains('年'))
                {
                    date = objs["tables_result"][0]["body"][55]["words"].ToString().Replace("\n", "").Replace("\r", "");
                }
                string date_or_month_pattern = @"\d{4}年\d{1,2}[\u4e00-\u9fa5]\d{0,}日|\d{4}年\d{1,2}[\u4e00-\u9fa5]";
                MatchCollection dateNeed = Regex.Matches(date, date_or_month_pattern);
                if (dateNeed.Count == 2)
                {
                    date = dateNeed[1].Value;
                    string date_or_month_pattern2 = @"\d+";
                    MatchCollection matches = Regex.Matches(date, date_or_month_pattern2);
                    int year = int.Parse(matches[0].Value);
                    int month = int.Parse(matches[1].Value);
                    int day = int.Parse(matches[2].Value);
                    date = matches[0].Value + "年" + matches[1].Value + "月" + matches[2].Value + "日";
                    DateTime dateforcell = new DateTime(year, month, day);
                    //计算2年后的日期
                    string nextdate;
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
            string speed;
            try
            {
                Console.WriteLine("输入限速器额定速度：");
                speed = Console.ReadLine();
                Console.WriteLine("速度：" + speed);
            }
            catch
            {
                Console.WriteLine("速度获取错误");
                speed = "/";
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
                xiansuqiDirection = "☑单向☐双向";
                xiansuqiDirectionForReport = "单向";
            }
            else
            {
                xiansuqiDirection = "☐单向☑双向";
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

            return result;
        }
    }
}
