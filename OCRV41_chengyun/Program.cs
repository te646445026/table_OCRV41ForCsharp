﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using RestSharp;//依赖版本106.15.0 https://www.nuget.org/packages/RestSharp/106.15.0
using Newtonsoft.Json; //https://www.nuget.org/packages/Newtonsoft.Json
using Newtonsoft.Json.Linq;
using System.Collections;
using NPOI.XWPF.UserModel;



namespace OCRV41_chengyun
{
    internal class Program
    {
        // KEY信息
        const string API_KEY = "Et4nGdx8ecc5chOnoilbxEyX";
        const string SECRET_KEY = "505cd0eiUZt22mPzelDGVrWzN7ELwteh";
        const string REQUEST_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/table";

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
            string? workPath;
            string data_dir = "";
            string folder_dir = "";

            ArrayList result_dir = new ArrayList();
            Dictionary<string, string> jsonMessage = new Dictionary<string, string>();

            PathMessage path;

            path = CheckDefaultPath();
            workPath = path.FolderPath;


            Console.WriteLine("已生成识别结果请按0，未生成请按1：");
            string? situation = Console.ReadLine();

            if (situation == "1")
            {
                //从json文件中读取

                data_dir = path.DataFilePath + "\\";

                folder_dir = path.DataJsonFilePath + "\\";


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
                    foreach (string fileName in fileDialog.FileNames)
                    {
                        result_dir.Add(fileName); // 获取用户选择的多个文件名的数组                                                              // 处理用户选择的文件路径
                    }
                }
            }
            if (situation == "1")
            {
                int num = 0;
                DirectoryInfo directoryInfo = new DirectoryInfo(data_dir);
                foreach (FileInfo file in directoryInfo.GetFiles())
                {
                    Console.WriteLine("{0}: {1} 正在处理：", num + 1, file.Name.Split('.')[0]);
                    string data_json = BaiduApi(file.FullName, REQUEST_URL, GetAccessToken());
                    string jsonFile_name = folder_dir + file.Name.Split('.')[0] + ".json";
                    File.WriteAllText(jsonFile_name, data_json);

                    Console.WriteLine("{0}: {1} 下载完成。", num + 1, jsonFile_name);
                    num++;
                    Console.WriteLine("--------------------------------------");
                    Console.WriteLine("");
                    Thread.Sleep(1000);
                }

            }
            if (situation == "1")
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
                //分别选中乘运质量软件生成的原始记录和报告，通过窗口选中确定路径

                //FileStream docFlieRec = new FileStream(workPath + "\\限速器测试记录模板3.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //FileStream docFlieRep = new FileStream(workPath + "\\限速器测试报告模板3.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite);

                //选择原始记录路径
                // 创建 OpenFileDialog 对象
                OpenFileDialog fileDialogRec = new OpenFileDialog();
                // 设置对话框的属性
                fileDialogRec.Multiselect = false; // 不允许多选文件
                fileDialogRec.Title = "请选择自动生成的原始记录"; // 设置对话框的标题
                fileDialogRec.Filter = "doc文件(*.doc)|*.doc"; // 设置对话框的文件过滤器

                // 显示对话框并获取用户选择的文件路径
                FileStream docFlieRec;
                DialogResult result = fileDialogRec.ShowDialog();
                if (result == DialogResult.OK)
                {
                   docFlieRec = new FileStream(fileDialogRec.FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                }
                else
                {
                    Console.WriteLine("没有选中文档，程序结束运行");
                    break;
                }

                //选择报告路径
                // 创建 OpenFileDialog 对象
                OpenFileDialog fileDialogRep = new OpenFileDialog();
                // 设置对话框的属性
                fileDialogRep.Multiselect = false; // 不允许多选文件
                fileDialogRep.Title = "请选择自动生成的报告"; // 设置对话框的标题
                fileDialogRep.Filter = "doc文件(*.doc)|*.doc"; // 设置对话框的文件过滤器

                // 显示对话框并获取用户选择的文件路径
                FileStream docFlieRep;
                result = fileDialogRep.ShowDialog();
                if (result == DialogResult.OK)
                {
                    docFlieRep = new FileStream(fileDialogRep.FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                }
                else
                {
                    Console.WriteLine("没有选中文档，程序结束运行");
                    break;
                }

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
                    paragraphsRec[0].CreateRun().SetText(jsonMessage["reportNum2"]);
                    //if (jsonMessage["xiansuqiDirectionForReport"] == "双向")
                    //{
                    //    paragraphsRec[0].CreateRun().SetText("D");
                    //}
                }
                catch
                {
                    Console.WriteLine("reportNum2 write error");
                }


                string outPath = string.Format(workPath + "\\{0}_{1}_{2}.docx",
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


                string outPath2 = string.Format(workPath + "\\{0}.docx",
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
                    path.DataFilePath = folder1.SelectedPath;
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine($"已选择需要识别的图片所在文件夹：{path.DataFilePath}");
                    Console.WriteLine("");
                }

                FolderBrowserDialog folder2 = new FolderBrowserDialog();
                folder2.Description = "请选择识别结果存放的文件夹";

                if (folder2.ShowDialog() == DialogResult.OK)
                {
                    path.DataJsonFilePath = folder2.SelectedPath;
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine($"已选择识别结果存放的文件夹：{path.DataJsonFilePath}");
                    Console.WriteLine("");
                }
                //确定工作路径

                path.FolderPath = Path.GetDirectoryName(folder1.SelectedPath);

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
                path.DataJsonFilePath = defaultStrPath.DataJsonFilePath;
                //确定工作路径
                path.FolderPath = Path.GetDirectoryName(path.DataFilePath);
                Console.WriteLine($"当前工作文件夹目录为：{path.FolderPath}");
            }

            return path;
        }

        static int ObjsIndex(string str, JObject objs)
        {
            int index = 0;
            for (int i = 0; i < objs["tables_result"][0]["body"].Count(); i++)
            {
                var isContain = objs["tables_result"][0]["body"][i]["words"].ToString().Contains(str);
                if (isContain)
                {
                    index = i;
                    break;
                }

            }
            return index;
        }

        static Dictionary<string, string> JsonMessage(string filePath)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            string json = File.ReadAllText(filePath);
            JObject? objs = JObject.Parse(json);

            string deviceCode;
            try
            {
                int index = ObjsIndex("设备代码", objs) + 1;

                deviceCode = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");

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
                int index = ObjsIndex("型号", objs) + 1;
                model = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                int index = ObjsIndex("产品编号", objs) + 1;
                serialNum = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                int index = ObjsIndex("制造单位", objs) + 1;
                ManufacturingUnit = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                int index = ObjsIndex("使用单位", objs) + 1;
                userName = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                int index = ObjsIndex("使用地点", objs) + 1;
                UsingAddress = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                int index = ObjsIndex("维护保养单位", objs) + 1;
                MaintenanceUnit = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                int index = ObjsIndex("额定速度", objs) + 1;
                speed = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "");
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
                int index = ObjsIndex("温度", objs);
                temperature = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");
                string temperature_pattern = @"\d{2,3}";
                MatchCollection temperatureNeed = Regex.Matches(temperature, temperature_pattern);
                temperature = $"温度：{temperatureNeed[0].ToString()}℃，  湿度：{temperatureNeed[1].ToString()}％ ， 电压：{temperatureNeed[2].ToString()}V";
                Console.WriteLine("温度、湿度、电压: " + temperature);
            }
            catch
            {
                temperature = "温度：    ℃，  湿度：    ％ ， 电压：     V";

                Console.WriteLine("温度、湿度、电压自动获取失败");
            }

            string reportNum;
            string reportNum2;
            try
            {
                int index = 0;
                for (int i = 0; i < objs["tables_result"][0]["header"].Count(); i++)
                {
                    var isContain = objs["tables_result"][0]["header"][i]["words"].ToString().Contains("编号");
                    if (isContain)
                    {
                        index = i;
                        break;
                    }

                }
                reportNum = objs["tables_result"][0]["header"][index]["words"].ToString();
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
            string? date;
            string next_year;
            string next_year_flag;
            string shenhe_date;
            try
            {
                int index = ObjsIndex("\n校核", objs);
                date = objs["tables_result"][0]["body"][index]["words"].ToString().Replace("\n", "").Replace("\r", "");
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
                    Console.WriteLine("检验时间为：" + date);
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

            return result;
        }
    }
}

