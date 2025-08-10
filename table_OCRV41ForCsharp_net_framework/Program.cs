using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NLog;
using NLog.Extensions.Logging;
using table_OCRV41ForCsharp_net_framework.Interfaces;
using table_OCRV41ForCsharp_net_framework.Models;
using table_OCRV41ForCsharp_net_framework.Services;

namespace table_OCRV41ForCsharp_net_framework
{
    internal class Program
    {
        private static ServiceProvider serviceProvider;
        private static ILogger<Program> logger;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // 初始化 NLog 配置
                var nlogConfig = new NLog.Config.XmlLoggingConfiguration("nlog.config");
                LogManager.Configuration = nlogConfig;
                
                // 配置服务
                ConfigureServices();
                logger = serviceProvider.GetService<ILogger<Program>>();
                logger.LogInformation("应用程序启动");

                // 显示欢迎界面
                DisplayWelcomeScreen();

                // 检查和初始化路径
                var pathService = serviceProvider.GetService<IPathService>();
                var pathMessage = pathService.CheckDefaultPath();

                // 检查密钥
                var keyService = serviceProvider.GetService<KeyService>();
                var key = keyService.CheckKey();

                // 显示主菜单
                DisplayMainMenu();
                string choice = Console.ReadLine();

                if (choice == "0")
                {
                    // 从JSON文件读取数据
                    ProcessFromJsonFiles(pathMessage);
                }
                else if (choice == "1")
                {
                    // OCR识别模式
                    ProcessFromOcrImages(pathMessage, key);
                }
                else
                {
                    Console.WriteLine("无效的选择，程序退出。");
                    return;
                }
            }
            catch (Exception ex)
            {
                logger?.LogError(ex, "程序执行过程中发生未处理的异常");
                MessageBox.Show($"程序执行过程中发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                logger?.LogInformation("应用程序结束");
                serviceProvider?.Dispose();
            }
        }

        /// <summary>
        /// 配置依赖注入服务
        /// </summary>
        private static void ConfigureServices()
        {
            var services = new ServiceCollection();

            // 配置日志 - 使用 NLog
            services.AddLogging(builder =>
            {
                builder.ClearProviders();
                builder.AddNLog();
                builder.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Information);
            });

            // 注册服务
            services.AddSingleton<IPathService, PathService>();
            services.AddSingleton<KeyService>();
            services.AddSingleton<GetFileContentAsBase64Service>();
            services.AddSingleton<TencentOcrParser>();
            services.AddSingleton<TencentOcrService>();
            services.AddSingleton<HolidayService>();

            serviceProvider = services.BuildServiceProvider();
        }

        /// <summary>
        /// 从JSON文件处理数据
        /// </summary>
        private static void ProcessFromJsonFiles(PathMessage pathMessage)
        {
            try
            {
                logger.LogInformation("开始从JSON文件处理数据");

                List<string> jsonFiles = new List<string>();
                
                // 直接打开文件选择对话框
                try
                {
                    logger.LogInformation("打开文件选择对话框");
                    OpenFileDialog fileDialog = new OpenFileDialog
                    {
                        Multiselect = true,
                        Title = "请选择文件",
                        Filter = "json文件(*.json)|*.json"
                    };

                    DialogResult result = fileDialog.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        jsonFiles.AddRange(fileDialog.FileNames);
                        foreach (string fileName in fileDialog.FileNames)
                        {
                            logger.LogInformation($"选择文件: {fileName}");
                        }
                    }
                    else
                    {
                        logger.LogWarning("用户取消了文件选择");
                        Console.WriteLine("未选择文件，程序退出。");
                        Console.ReadKey();
                        return;
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "打开文件选择对话框时出错");
                    MessageBox.Show($"选择文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (jsonFiles.Count == 0)
                {
                    Console.WriteLine("没有找到或选择JSON文件，程序退出。");
                    Console.ReadKey();
                    return;
                }

                Console.WriteLine($"找到 {jsonFiles.Count} 个JSON文件");
                string workPath = pathMessage.FolderPath;

                foreach (string jsonPath in jsonFiles)
                {
                    try
                    {
                        logger.LogInformation($"处理JSON文件: {jsonPath}");
                        Console.WriteLine($"正在处理: {Path.GetFileName(jsonPath)}");

                        string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                        var tencentOcrParser = serviceProvider.GetService<TencentOcrParser>();
                        var resultForJsonMessage = tencentOcrParser.Parse(jsonContent);

                        if (resultForJsonMessage == null)
                        {
                            logger.LogWarning($"JSON文件 {jsonPath} 解析失败");
                            continue;
                        }

                        // 生成Word文档
                        GenerateWordDocuments(resultForJsonMessage, workPath, jsonPath);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, $"处理JSON文件 {jsonPath} 时出错");
                        MessageBox.Show($"处理文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                logger.LogInformation("所有文件处理完成");
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("🎉 处理完成！按任意键退出程序");
                Console.ResetColor();
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "从JSON文件处理数据时发生错误");
                throw;
            }
        }

        /// <summary>
        /// 从OCR图片处理数据
        /// </summary>
        private static void ProcessFromOcrImages(PathMessage pathMessage, KEY key)
        {
            try
            {
                logger.LogInformation("开始OCR图片识别处理");

                var getFileContentAsBase64Service = serviceProvider.GetService<GetFileContentAsBase64Service>();
                var tencentOcrService = serviceProvider.GetService<TencentOcrService>();
                var tencentOcrParser = serviceProvider.GetService<TencentOcrParser>();

                // 选择图片文件
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Title = "选择要识别的图片文件",
                    Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif|所有文件|*.*",
                    Multiselect = true
                };

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    Console.WriteLine("未选择文件，程序退出。");
                    return;
                }

                string workPath = pathMessage.FolderPath;
                string[] selectedFiles = openFileDialog.FileNames;

                Console.WriteLine($"选择了 {selectedFiles.Length} 个文件进行处理");

                foreach (string filePath in selectedFiles)
                {
                    try
                    {
                        logger.LogInformation($"处理图片文件: {filePath}");
                        Console.WriteLine($"正在处理: {Path.GetFileName(filePath)}");

                        // 转换为Base64
                        string base64Content = getFileContentAsBase64Service.GetFileContentAsBase64(filePath);

                        // OCR识别
                        string ocrResponse = tencentOcrService.RecognizeTable(base64Content, key);

                        // 解析OCR结果
                        var ocrResult = tencentOcrParser.Parse(ocrResponse);

                        // 保存JSON结果
                        string jsonFileName = Path.GetFileNameWithoutExtension(filePath) + ".json";
                        string jsonPath = Path.Combine(pathMessage.DataFilePath, jsonFileName);
                        string jsonContent = JsonConvert.SerializeObject(ocrResult, Formatting.Indented);
                        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

                        logger.LogInformation($"OCR结果已保存到: {jsonPath}");

                        // 生成Word文档
                        GenerateWordDocuments(ocrResult, workPath, jsonPath);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, $"处理图片文件 {filePath} 时出错");
                        MessageBox.Show($"处理文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                logger.LogInformation("所有文件处理完成");
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("🎉 处理完成！按任意键退出程序");
                Console.ResetColor();
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "OCR图片处理时发生错误");
                throw;
            }
        }

        /// <summary>
        /// 生成Word文档
        /// </summary>
        private static void GenerateWordDocuments(OcrResult resultForJsonMessage, string workPath, string jsonPath)
        {
            try
            {
                string recordTemplatePath = workPath + "\\限速器测试记录模板4.docx";
                string reportTemplatePath = workPath + "\\限速器测试报告模板4.docx";

                if (!File.Exists(recordTemplatePath) || !File.Exists(reportTemplatePath))
                {
                    throw new FileNotFoundException("模板文件不存在", !File.Exists(recordTemplatePath) ? recordTemplatePath : reportTemplatePath);
                }

                logger.LogInformation("打开Word模板文件");
                FileStream docFlieRec = new FileStream(recordTemplatePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                FileStream docFlieRep = new FileStream(reportTemplatePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                XWPFDocument documentRec = new XWPFDocument(docFlieRec);
                XWPFDocument documentRep = new XWPFDocument(docFlieRep);

                IList<XWPFParagraph> paragraphsRec = documentRec.Paragraphs;
                Console.WriteLine(paragraphsRec[2].ParagraphText + resultForJsonMessage.ReportNum);

                IList<XWPFTable> tablesRec = documentRec.Tables;
                XWPFTable tableRec0 = tablesRec[0];
                XWPFTable tableRec1 = tablesRec[1];

                IList<XWPFParagraph> paragraphsRep = documentRep.Paragraphs;
                Console.WriteLine(paragraphsRep[3].ParagraphText + resultForJsonMessage.ReportNum);

                IList<XWPFTable> tablesRep = documentRep.Tables;
                XWPFTable tableRep0 = tablesRep[0];
                XWPFTable tableRep1 = tablesRep[1];

                logger.LogInformation("开始填充Word文档内容");

                // 填充记录模板
                FillRecordTemplate(tableRec0, paragraphsRec, resultForJsonMessage);

                // 保存记录文件
                string outPath = string.Format(workPath + "\\{0}_{1}_{2}.docx",
                                                    resultForJsonMessage.DeviceCode,
                                                    Path.GetFileNameWithoutExtension(jsonPath),
                                                    resultForJsonMessage.NextYearFlag);
                FileStream outFile = new FileStream(outPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                documentRec.Write(outFile);
                outFile.Close();
                documentRec.Close();
                docFlieRec.Close();

                Console.WriteLine("{0}打印记录完成", Path.GetFileNameWithoutExtension(jsonPath));
                Console.WriteLine("-------------------------------------------------------");
                Console.WriteLine("");

                // 填充报告模板
                FillReportTemplate(tableRep0, paragraphsRep, resultForJsonMessage);

                // 保存报告文件
                string outPath2 = string.Format(workPath + "\\{0}.docx", resultForJsonMessage.DeviceCode);
                logger.LogInformation($"保存报告文件: {outPath2}");
                FileStream outFile2 = new FileStream(outPath2, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                documentRep.Write(outFile2);
                outFile2.Close();
                documentRep.Close();
                docFlieRep.Close();

                Console.WriteLine("{0}打印报告完成", Path.GetFileNameWithoutExtension(jsonPath));
                Console.WriteLine("-------------------------------------------------------");
                Console.WriteLine("");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "生成Word文档时发生错误");
                throw;
            }
        }

        /// <summary>
        /// 填充记录模板
        /// </summary>
        private static void FillRecordTemplate(XWPFTable tableRec0, IList<XWPFParagraph> paragraphsRec, OcrResult resultForJsonMessage)
        {
            try
            {
                tableRec0.GetRow(0).GetCell(1).SetText(resultForJsonMessage.UserName);
                tableRec0.GetRow(0).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "设置委托单位时出错");
                Console.WriteLine("userName write error");
            }

            try
            {
                tableRec0.GetRow(1).GetCell(1).SetText(resultForJsonMessage.UserName);
                tableRec0.GetRow(1).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "设置使用单位时出错");
                Console.WriteLine("userName write error");
            }

            try
            {
                tableRec0.GetRow(2).GetCell(1).SetText(resultForJsonMessage.MaintenanceUnit);
                tableRec0.GetRow(2).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("MaintenanceUnit write error");
            }

            try
            {
                tableRec0.GetRow(3).GetCell(1).SetText(resultForJsonMessage.UsingAddress);
                tableRec0.GetRow(3).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("UsingAddress write error");
            }

            try
            {
                tableRec0.GetRow(4).GetCell(2).SetText(resultForJsonMessage.ElevatorDeviceType);
                tableRec0.GetRow(4).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("ElevatorDeviceType write error");
            }

            try
            {
                tableRec0.GetRow(4).GetCell(4).SetText(resultForJsonMessage.DeviceCode);
                tableRec0.GetRow(4).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("DeviceCode write error");
            }

            try
            {
                tableRec0.GetRow(5).GetCell(2).SetText(resultForJsonMessage.SerialNum);
                tableRec0.GetRow(5).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("SerialNum write error");
            }

            try
            {
                tableRec0.GetRow(5).GetCell(4).SetText(resultForJsonMessage.Speed + "m/s");
                tableRec0.GetRow(5).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("speed write error");
            }

            try
            {
                tableRec0.GetRow(6).GetCell(2).SetText(resultForJsonMessage.XiansuqiManufacturingUnit);
                tableRec0.GetRow(6).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiManufacturingUnit write error");
            }

            try
            {
                tableRec0.GetRow(7).GetCell(2).SetText(resultForJsonMessage.XiansuqiModel);
                tableRec0.GetRow(7).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiModel write error");
            }

            try
            {
                tableRec0.GetRow(7).GetCell(4).SetText(resultForJsonMessage.XiansuqiNum);
                tableRec0.GetRow(7).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiNum write error");
            }

            try
            {
                tableRec0.GetRow(8).GetCell(4).SetText(resultForJsonMessage.XiansuqiDirection);
                tableRec0.GetRow(8).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiDirection write error");
            }

            try
            {
                tableRec0.GetRow(10).GetCell(1).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.XiansuqiElectricalUpSpeed + "m/s");
                tableRec0.GetRow(10).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiElectricalUpSpeed write error");
            }

            try
            {
                tableRec0.GetRow(10).GetCell(2).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.XiansuqiElectricalDownSpeed + "m/s");
                tableRec0.GetRow(10).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiElectricalDownSpeed write error");
            }

            try
            {
                tableRec0.GetRow(10).GetCell(3).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.XiansuqiMechanicalUpSpeed + "m/s");
                tableRec0.GetRow(10).GetCell(3).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiMechanicalUpSpeed write error");
            }

            try
            {
                tableRec0.GetRow(10).GetCell(4).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.XiansuqiMechanicalDownSpeed + "m/s");
                tableRec0.GetRow(10).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiMechanicalDownSpeed write error");
            }

            try
            {
                tableRec0.GetRow(25).GetCell(0).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.Date);
                tableRec0.GetRow(25).GetCell(0).Paragraphs[0].Alignment = ParagraphAlignment.RIGHT;
            }
            catch
            {
                Console.WriteLine("date write error");
            }

            try
            {
                paragraphsRec[2].CreateRun().SetText(resultForJsonMessage.JianyanOrjiance?.Equals("检验") == true ? "D" : "E");
                paragraphsRec[2].CreateRun().SetText(resultForJsonMessage.ReportNum);
                paragraphsRec[2].Alignment = ParagraphAlignment.RIGHT;
            }
            catch
            {
                Console.WriteLine("reportNum2 write error");
            }
        }

        /// <summary>
        /// 填充报告模板
        /// </summary>
        private static void FillReportTemplate(XWPFTable tableRep0, IList<XWPFParagraph> paragraphsRep, OcrResult resultForJsonMessage)
        {
            try
            {
                tableRep0.GetRow(0).GetCell(1).SetText(resultForJsonMessage.UserName);
                tableRep0.GetRow(0).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("userName write error");
            }

            try
            {
                tableRep0.GetRow(1).GetCell(1).SetText(resultForJsonMessage.UserName);
                tableRep0.GetRow(1).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("userName write error");
            }

            try
            {
                tableRep0.GetRow(2).GetCell(1).SetText(resultForJsonMessage.MaintenanceUnit);
                tableRep0.GetRow(2).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("MaintenanceUnit write error");
            }

            try
            {
                tableRep0.GetRow(3).GetCell(1).SetText(resultForJsonMessage.UsingAddress);
                tableRep0.GetRow(3).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("UsingAddress write error");
            }

            try
            {
                tableRep0.GetRow(4).GetCell(2).SetText(resultForJsonMessage.ElevatorDeviceType);
                tableRep0.GetRow(4).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("ElevatorDeviceType write error");
            }

            try
            {
                tableRep0.GetRow(4).GetCell(4).SetText(resultForJsonMessage.DeviceCode);
                tableRep0.GetRow(4).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("DeviceCode write error");
            }

            try
            {
                tableRep0.GetRow(5).GetCell(2).SetText(resultForJsonMessage.SerialNum);
                tableRep0.GetRow(5).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("SerialNum write error");
            }

            try
            {
                tableRep0.GetRow(5).GetCell(4).SetText(resultForJsonMessage.Speed + "m/s");
                tableRep0.GetRow(5).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("Speed write error");
            }

            try
            {
                tableRep0.GetRow(6).GetCell(2).SetText(resultForJsonMessage.XiansuqiManufacturingUnit);
                tableRep0.GetRow(6).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiManufacturingUnit write error");
            }

            try
            {
                tableRep0.GetRow(7).GetCell(2).SetText(resultForJsonMessage.XiansuqiModel);
                tableRep0.GetRow(7).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiModel write error");
            }

            try
            {
                tableRep0.GetRow(7).GetCell(4).SetText(resultForJsonMessage.XiansuqiNum);
                tableRep0.GetRow(7).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiNum write error");
            }


            try
            {
                tableRep0.GetRow(8).GetCell(4).SetText(resultForJsonMessage.XiansuqiDirection);
                tableRep0.GetRow(8).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiDirection write error");
            }

            try
            {
                tableRep0.GetRow(10).GetCell(1).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.XiansuqiElectricalUpSpeed + "m/s");
                tableRep0.GetRow(10).GetCell(1).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiElectricalUpSpeed write error");
            }

            try
            {
                tableRep0.GetRow(10).GetCell(2).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.XiansuqiElectricalDownSpeed + "m/s");
                tableRep0.GetRow(10).GetCell(2).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiElectricalDownSpeed write error");
            }

            try
            {
                tableRep0.GetRow(10).GetCell(3).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.XiansuqiMechanicalUpSpeed + "m/s");
                tableRep0.GetRow(10).GetCell(3).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiMechanicalUpSpeed write error");
            }

            try
            {
                tableRep0.GetRow(10).GetCell(4).Paragraphs[0].CreateRun().SetText(resultForJsonMessage.XiansuqiMechanicalDownSpeed + "m/s");
                tableRep0.GetRow(10).GetCell(4).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("XiansuqiMechanicalDownSpeed write error");
            }

            try
            {
                tableRep0.GetRow(22).GetCell(0).SetText(resultForJsonMessage.Date);
                tableRep0.GetRow(22).GetCell(0).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("Date write error");
            }

            try
            {
                tableRep0.GetRow(23).GetCell(0).SetText(resultForJsonMessage.ShenheDate);
                tableRep0.GetRow(23).GetCell(0).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("ShenheDate write error");
            }

            try
            {
                tableRep0.GetRow(24).GetCell(0).SetText(resultForJsonMessage.ShenheDate);
                tableRep0.GetRow(24).GetCell(0).Paragraphs[0].Alignment = ParagraphAlignment.LEFT;
            }
            catch
            {
                Console.WriteLine("ShenheDate write error");
            }

            try
            {
                paragraphsRep[3].CreateRun().SetText(resultForJsonMessage.JianyanOrjiance?.Equals("检验") == true ? "D" : "E");
                paragraphsRep[3].CreateRun().SetText(resultForJsonMessage.ReportNum);
                paragraphsRep[3].Alignment = ParagraphAlignment.RIGHT;

                // 检查段落是否已有Run，如果有则复制格式
                var newRun = paragraphsRep[15].CreateRun();
                if (paragraphsRep[15].Runs.Count > 1)
                {
                    var existingRun = paragraphsRep[15].Runs[0];
                    // 复制字体格式
                    newRun.FontSize = existingRun.FontSize;
                    newRun.FontFamily = existingRun.FontFamily;
                    newRun.IsBold = existingRun.IsBold;
                    newRun.IsItalic = existingRun.IsItalic;
                    newRun.Underline = UnderlinePatterns.Single;
                }
                else
                {
                    newRun.Underline = UnderlinePatterns.Single; // 设置下划线
                }
                newRun.SetText(resultForJsonMessage.UserName);

                newRun = paragraphsRep[17].CreateRun();
                if (paragraphsRep[17].Runs.Count > 1)
                {
                    var existingRun = paragraphsRep[17].Runs[0];
                    // 复制字体格式
                    newRun.FontSize = existingRun.FontSize;
                    newRun.FontFamily = existingRun.FontFamily;
                    newRun.IsBold = existingRun.IsBold;
                    newRun.IsItalic = existingRun.IsItalic;
                    newRun.Underline = UnderlinePatterns.Single;
                }
                else
                {
                    newRun.Underline = UnderlinePatterns.Single; // 设置下划线
                }
                newRun.SetText(resultForJsonMessage.Date);

                paragraphsRep[53].CreateRun().SetText(resultForJsonMessage.JianyanOrjiance?.Equals("检验") == true ? "D" : "E");
                paragraphsRep[53].CreateRun().SetText(resultForJsonMessage.ReportNum);
                paragraphsRep[53].Alignment = ParagraphAlignment.RIGHT;
            }
            catch
            {
                Console.WriteLine("reportNum2 write error");
            }
        }

        /// <summary>
        /// 显示程序启动欢迎界面
        /// </summary>
        private static void DisplayWelcomeScreen()
         {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║                    电梯限速器检测报告生成系统                    ║");
            Console.WriteLine("║                    Elevator Speed Limiter Report              ║");
            Console.WriteLine("║                         Generation System                     ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════╝");
            Console.ResetColor();
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("📋 功能说明:");
            Console.WriteLine("   • 支持OCR图片识别，自动提取检测数据");
            Console.WriteLine("   • 支持从JSON文件读取已识别的数据");
            Console.WriteLine("   • 自动生成标准格式的Word检测报告");
            Console.ResetColor();
            Console.WriteLine();
        }

        /// <summary>
        /// 显示主菜单选择界面
        /// </summary>
        private static void DisplayMainMenu()
        {
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("┌─────────────────────────────────────────┐");
            Console.WriteLine("│                 选择操作模式                │");
            Console.WriteLine("└─────────────────────────────────────────┘");
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine("📂 [0] 使用已生成的识别结果文件");
            Console.WriteLine("🖼️ [1] 上传图片进行OCR识别");
            Console.WriteLine();
            Console.Write("请输入您的选择 [0/1]: ");
        }
    }
}