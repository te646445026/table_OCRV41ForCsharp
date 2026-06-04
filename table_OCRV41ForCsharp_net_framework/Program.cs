using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;
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
                var keyService = serviceProvider.GetService<IKeyService>();
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
                        ProcessFromOcrImages(serviceProvider);
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

            // 配置 NLog
            LogManager.LoadConfiguration("nlog.config");
            
            // 配置日志 - 使用 NLog
            services.AddLogging(builder =>
            {
                builder.ClearProviders();
                builder.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Information);
                builder.AddNLog();
            });
            
            // 添加配置服务
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("default.json", optional: true, reloadOnChange: true)
                .Build();
            services.AddSingleton<IConfiguration>(configuration);
            
            // 注册业务服务
            services.AddSingleton<IPathService, PathService>();
            services.AddSingleton<IKeyService, KeyService>();
            services.AddSingleton<IGetFileContentAsBase64Service, GetFileContentAsBase64Service>();
            services.AddSingleton<IOcrParser, TencentOcrParser>();
            services.AddSingleton<ISpeedValueGenerator, SpeedValueGenerator>();
            services.AddSingleton<IWordTemplateFiller, WordTemplateFiller>();

            // 使用工厂模式注册 TencentOcrService
            services.AddSingleton<IOcrService>(provider =>
            {
                var keyService = provider.GetService<IKeyService>();
                var secretId = keyService.CheckKey().API_KEY;
                var secretKey = keyService.CheckKey().SECRET_KEY;
                return new TencentOcrService(secretId, secretKey);
            });
            services.AddSingleton<HolidayService>();

            serviceProvider = services.BuildServiceProvider();
        }

        /// <summary>
        /// 初始化全局异常处理
        /// </summary>
        /// <param name="logger">日志记录器</param>
        private static void InitializeExceptionHandler(ILogger<Program> logger)
        {
            AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
            {
                var exception = (Exception)args.ExceptionObject;
                logger?.LogCritical(exception, "程序遇到了未处理的异常");
                MessageBox.Show("程序遇到了未处理的异常，请查看日志文件获取详细信息。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
        }

        /// <summary>
        /// 从JSON文件处理数据
        /// </summary>
        private static void ProcessFromJsonFiles(PathMessage pathMessage)
        {
            // 获取日志记录器
            var logger = serviceProvider.GetService<ILogger<Program>>();
            
            // 初始化异常处理
            InitializeExceptionHandler(logger);

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
                        Filter = "json文件(*.json)|*.json",
                        InitialDirectory = pathMessage.DataJsonFilePath
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
                        var ocrParser = serviceProvider.GetService<IOcrParser>();
                        var resultForJsonMessage = ocrParser.Parse(jsonContent);

                        if (resultForJsonMessage == null)
                        {
                            logger.LogWarning($"JSON文件 {jsonPath} 解析失败");
                            continue;
                        }

                        // 生成Word文档
                        GenerateWordDocuments(serviceProvider, resultForJsonMessage, workPath, jsonPath);
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
        /// 从图片识别处理
        /// </summary>
        /// <param name="serviceProvider">服务提供者</param>
        private static void ProcessFromOcrImages(IServiceProvider serviceProvider)
        {
            // 获取日志记录器
            var logger = serviceProvider.GetService<ILogger<Program>>();
            
            // 初始化异常处理
            InitializeExceptionHandler(logger);

            try
            {
                logger?.LogInformation("开始从图片识别处理");
                
                // 检查默认路径
                var pathService = serviceProvider.GetService<IPathService>();
                if (pathService == null)
                {
                    throw new Exception("无法获取路径服务");
                }
                
                PathMessage path = pathService.CheckDefaultPath();
                string workPath = path.FolderPath;
                logger?.LogInformation("工作路径: {WorkPath}", workPath);
                
                string dataDir = path.DataFilePath + "\\";
                string folderDir = path.DataJsonFilePath + "\\";
                
                logger?.LogInformation("数据目录: {DataDir}", dataDir);
                logger?.LogInformation("JSON文件目录: {FolderDir}", folderDir);
                
                // 获取服务
                var ocrService = serviceProvider.GetService<IOcrService>();
                var getFileContentAsBase64Service = serviceProvider.GetService<IGetFileContentAsBase64Service>();
                var ocrParser = serviceProvider.GetService<IOcrParser>();
                
                if (ocrService == null || getFileContentAsBase64Service == null || ocrParser == null)
                {
                    throw new Exception("无法解析所需的服务");
                }
                
                // 处理图片文件
                logger?.LogInformation("开始处理图片文件");
                int num = 0;
                DirectoryInfo directoryInfo = new DirectoryInfo(dataDir);

                if (!directoryInfo.Exists)
                {
                    throw new DirectoryNotFoundException($"目录不存在: {dataDir}");
                }

                foreach (FileInfo file in directoryInfo.GetFiles())
                {
                    try
                    {
                        logger?.LogInformation("处理文件: {FileName}", file.Name);
                        Console.WriteLine("{0}: {1} 正在处理：", num + 1, file.Name.Split('.')[0]);

                        // 获取KEY对象
                        var keyService = serviceProvider.GetService<IKeyService>();
                        var key = keyService.CheckKey();
                        
                        string imageBase64 = getFileContentAsBase64Service.GetFileContentAsBase64(file.FullName);
                        string data_json = ocrService.RecognizeTable(imageBase64, key);
                        string jsonFile_name = folderDir + file.Name.Split('.')[0] + ".json";

                        File.WriteAllText(jsonFile_name, data_json);
                        logger?.LogInformation("文件处理完成: {JsonFileName}", jsonFile_name);

                        Console.WriteLine("{0}: {1} 下载完成。", num + 1, jsonFile_name);
                        num++;
                        Console.WriteLine("--------------------------------------");
                        Console.WriteLine("");
                        Thread.Sleep(1000);
                    }
                    catch (Exception ex)
                    {
                        logger?.LogError(ex, "处理文件 {FileName} 时出错", file.Name);
                        Console.WriteLine($"处理文件 {file.Name} 时出错: {ex.Message}");
                    }
                }
                
                // 生成Word文档
                string[] jsonFiles = Directory.GetFiles(folderDir, "*.json");
                
                foreach (string jsonPath in jsonFiles)
                {
                    try
                    {
                        logger.LogInformation($"处理JSON文件: {jsonPath}");
                        Console.WriteLine($"正在处理: {Path.GetFileName(jsonPath)}");

                        string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                        var resultForJsonMessage = ocrParser.Parse(jsonContent);

                        if (resultForJsonMessage == null)
                        {
                            logger.LogWarning($"JSON文件 {jsonPath} 解析失败");
                            continue;
                        }

                        // 生成Word文档
                        GenerateWordDocuments(serviceProvider, resultForJsonMessage, workPath, jsonPath);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, $"处理JSON文件 {jsonPath} 时出错");
                        MessageBox.Show($"处理文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                
                logger?.LogInformation("从图片识别处理完成");
                Console.WriteLine("处理完成！");
            }
            catch (Exception ex)
            {
                logger?.LogError(ex, "从图片识别处理时出错");
                MessageBox.Show($"处理过程中发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 生成Word文档
        /// </summary>
        private static void GenerateWordDocuments(IServiceProvider serviceProvider, OcrResult resultForJsonMessage, string workPath, string jsonPath)
        {
            var wordFiller = serviceProvider.GetService<IWordTemplateFiller>();

            string recordTemplatePath = Path.Combine(workPath, "限速器测试记录模板5.docx");
            string reportTemplatePath = Path.Combine(workPath, "限速器测试报告模板5.docx");

            if (!File.Exists(recordTemplatePath) || !File.Exists(reportTemplatePath))
            {
                throw new FileNotFoundException("模板文件不存在，请确保模板文件限速器测试记录模板5.docx和限速器测试报告模板5.docx存在于工作目录中", 
                    !File.Exists(recordTemplatePath) ? recordTemplatePath : reportTemplatePath);
            }

            logger.LogInformation("打开Word模板文件");

            using (var docStreamRec = new FileStream(recordTemplatePath, FileMode.Open, FileAccess.Read))
            using (var documentRec = new XWPFDocument(docStreamRec))
            {
                wordFiller.FillTemplate(documentRec, resultForJsonMessage);

                string outPath = Path.Combine(workPath, $"{resultForJsonMessage.DeviceCode}_{Path.GetFileNameWithoutExtension(jsonPath)}_{resultForJsonMessage.NextYearFlag}.docx");
                using (var outFile = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                {
                    documentRec.Write(outFile);
                }

                Console.WriteLine("{0}打印记录完成", Path.GetFileNameWithoutExtension(jsonPath));
                Console.WriteLine("-------------------------------------------------------");
                Console.WriteLine("");
            }

            using (var docStreamRep = new FileStream(reportTemplatePath, FileMode.Open, FileAccess.Read))
            using (var documentRep = new XWPFDocument(docStreamRep))
            {
                wordFiller.FillTemplate(documentRep, resultForJsonMessage);

                string outPath2 = Path.Combine(workPath, $"{resultForJsonMessage.DeviceCode}.docx");
                logger.LogInformation($"保存报告文件: {outPath2}");
                using (var outFile2 = new FileStream(outPath2, FileMode.Create, FileAccess.Write))
                {
                    documentRep.Write(outFile2);
                }

                Console.WriteLine("{0}打印报告完成", Path.GetFileNameWithoutExtension(jsonPath));
                Console.WriteLine("-------------------------------------------------------");
                Console.WriteLine("");
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