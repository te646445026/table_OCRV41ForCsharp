using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using table_OCRV41ForCsharp_net_framework.Interfaces;
using table_OCRV41ForCsharp_net_framework.Models;

namespace table_OCRV41ForCsharp_net_framework.Services
{
    public class TencentOcrParser : IOcrParser
    {
        public OcrResult Parse(string json)
        {
            var objs = JObject.Parse(json);
            OcrResult result = new OcrResult();

            // ==================== OCR识别结果解析 ====================
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("┌─────────────────────────────────────────┐");
            Console.WriteLine("│              OCR识别结果解析              │");
            Console.WriteLine("└─────────────────────────────────────────┘");
            Console.ResetColor();
            Console.WriteLine();

            result.JianyanOrjiance = "检测";
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("RTD", objs, out indexj, out indexi, out isContain);

                if (isContain)
                {
                    result.JianyanOrjiance = "检验";
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 检测类型: {result.JianyanOrjiance}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 获取检验还是检测失败，默认设置为检测");
                Console.ResetColor();
            }
            
            // 电梯设备品种 - 从OCR获取
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("设备品种", objs, out indexj, out indexi, out isContain);
                result.ElevatorDeviceType = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 电梯设备品种: {result.ElevatorDeviceType}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 电梯设备品种获取错误，已设置为默认值");
                Console.ResetColor();
                result.ElevatorDeviceType = "/";
            }

            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("设备代码", objs, out indexj, out indexi, out isContain);

                result.DeviceCode = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");

                Console.WriteLine("设备代码: " + result.DeviceCode);
            }
            catch
            {
                Console.WriteLine("设备代码获取错误");
                result.DeviceCode = "/";
            }

            //string model;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("产品型号", objs, out indexj, out indexi, out isContain);
                result.Model = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("产品型号: " + result.Model);
            }
            catch
            {
                Console.WriteLine("产品型号获取错误");
                result.Model = "/";
            }

            //string serialNum;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("产品编号", objs, out indexj, out indexi, out isContain);
                result.SerialNum = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("产品编号: " + result.SerialNum);
            }
            catch
            {
                Console.WriteLine("产品编号获取错误");
                result.SerialNum = "/";
            }

            //string ManufacturingUnit;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("制造单位名称", objs, out indexj, out indexi, out isContain);
                result.ManufacturingUnit = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("制造单位: " + result.ManufacturingUnit);
            }
            catch
            {
                Console.WriteLine("制造单位获取错误");
                result.ManufacturingUnit = "/";
            }

            //string userName;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("使用单位名称", objs, out indexj, out indexi, out isContain);
                result.UserName = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("使用单位: " + result.UserName);
            }
            catch
            {
                Console.WriteLine("使用单位获取错误");
                result.UserName = "/";
            }

            //string UsingAddress;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("安装地点", objs, out indexj, out indexi, out isContain);
                result.UsingAddress = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("安装地点: " + result.UsingAddress);
            }
            catch
            {
                Console.WriteLine("安装地点获取错误");
                result.UsingAddress = "/";
            }

            //string MaintenanceUnit;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("维护保养单位名称", objs, out indexj, out indexi, out isContain);
                result.MaintenanceUnit = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                Console.WriteLine("维护保养单位: " + result.MaintenanceUnit);
            }
            catch
            {
                Console.WriteLine("维护保养单位获取错误");
                result.MaintenanceUnit = "/";
            }
            
            //string speed;
            try
            {
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex("额定速度", objs, out indexj, out indexi, out isContain);
                result.Speed = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "");
                string speed_pattern = @"(\d+(\.\d+)?)";
                var speedNeed = Regex.Matches(result.Speed, speed_pattern);
                result.Speed = speedNeed[0].ToString();
                Console.WriteLine("速度：" + result.Speed);
            }
            catch
            {
                Console.WriteLine("速度获取错误");
                result.Speed = "/";
            }

            //string reportNum;
            //string reportNum2;
            string jianyanOrjianceReportNum;
            try
            {
                if (result.JianyanOrjiance.Equals("检验"))
                {
                    jianyanOrjianceReportNum = "RTD";
                }
                else
                {
                    jianyanOrjianceReportNum = "RTC";
                }
                int indexj;
                int indexi;
                bool isContain;
                ObjsIndex(jianyanOrjianceReportNum, objs, out indexj, out indexi, out isContain);

                result.ReportNum = objs["Response"]["TableDetections"][indexj]["Cells"][indexi]["Text"].ToString();
                //MatchCollection matchs = Regex.Matches(reportNum, @"^\d{8}");
                //reportNum2 = matchs[0].ToString().Substring(1,7);
                result.ReportNum = result.ReportNum.Substring(result.ReportNum.Length - 7);
                Console.WriteLine("报告编号: " + result.ReportNum);

            }
            catch
            {
                Console.WriteLine("报告编号获取错误");
                result.ReportNum = "/";
                
            }
            
            //string? date;
            //string next_year;
            //string next_year_flag;
            //string shenhe_date;
            string jianyanOrjianceDate;
            try
            {
                if (result.JianyanOrjiance.Equals("检验"))
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
                result.Date = objs["Response"]["TableDetections"][indexj]["Cells"][indexi]["Text"].ToString().Replace("\n", "").Replace("\r", "");
                string date_or_month_pattern = @"\d{4}年\d{1,2}[\u4e00-\u9fa5]\d{0,}日|\d{4}年\d{1,2}[\u4e00-\u9fa5]";
                MatchCollection dateNeed = Regex.Matches(result.Date, date_or_month_pattern);
                if (dateNeed != null && dateNeed.Count > 0)
                {
                    result.Date = dateNeed[dateNeed.Count - 1].Value;
                    string date_or_month_pattern2 = @"\d+";
                    MatchCollection matches = Regex.Matches(result.Date, date_or_month_pattern2);
                    int year = int.Parse(matches[0].Value);
                    int month = int.Parse(matches[1].Value);
                    int day = int.Parse(matches[2].Value);
                    result.Date = matches[0].Value + "年" + matches[1].Value + "月" + matches[2].Value + "日";
                    Console.WriteLine("检验时间为：" + result.Date);
                    DateTime dateforcell = new DateTime(year, month, day);
                    //计算2年后的日期（默认2年间隔）
                    DateTime next_year_date = dateforcell.AddYears(2);
                    result.NextYear = next_year_date.ToString("yyyy年MM月dd日");
                    result.NextYearFlag = "";
                    //计算审核校准日期（智能跳过节假日）
                    try
                    {
                        // 尝试使用异步API获取下一个工作日（跳过节假日和周末）
                        // 由于Parse方法不是异步的，我们需要使用Wait来等待异步操作完成
                        Task<DateTime> nextWorkingDayTask = HolidayService.GetNextWorkingDayAsync(dateforcell);
                        nextWorkingDayTask.Wait();
                        DateTime shenhe_dateforcell = nextWorkingDayTask.Result;
                        result.ShenheDate = shenhe_dateforcell.ToString("yyyy年MM月dd日");
                        Console.WriteLine("审核校准日期（已跳过节假日和周末）: " + result.ShenheDate);
                    }
                    catch (Exception ex)
                    {
                        // 如果API调用失败，回退到本地计算方法（只跳过周末）
                        Console.WriteLine($"使用API获取工作日失败，回退到本地计算: {ex.Message}");
                        DateTime shenhe_dateforcell = HolidayService.GetNextWorkingDay(dateforcell);
                        result.ShenheDate = shenhe_dateforcell.ToString("yyyy年MM月dd日");
                        Console.WriteLine("审核校准日期（已跳过周末）: " + result.ShenheDate);
                    }
                }
                else
                {
                    result.Date = "   年   月   日";
                    result.NextYear = "   年   月   日";
                    result.NextYearFlag = "检验日期和下检日期出错";
                    result.ShenheDate = "   年   月   日";
                    Console.WriteLine("检验日期获取错误");
                }
            }
            catch
            {
                Console.WriteLine("检验日期获取错误");
                result.Date = "   年   月   日";
                result.NextYear = "   年   月   日";
                result.NextYearFlag = "检验日期和下检日期出错";
                result.ShenheDate = "   年   月   日";
            }

             // 限速器制造单位 - 手动输入
            // ==================== 限速器基本信息录入 ====================
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("┌─────────────────────────────────────────┐");
            Console.WriteLine("│              限速器基本信息录入              │");
            Console.WriteLine("└─────────────────────────────────────────┘");
            Console.ResetColor();
            Console.WriteLine();

            try
            {
                Console.Write("请输入限速器制造单位: ");
                result.XiansuqiManufacturingUnit = Console.ReadLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 限速器制造单位: {result.XiansuqiManufacturingUnit}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 限速器制造单位获取错误，已设置为默认值");
                Console.ResetColor();
                result.XiansuqiManufacturingUnit = "/";
            }
            
            try
            {
                Console.Write("请输入限速器型号: ");
                result.XiansuqiModel = Console.ReadLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 限速器型号: {result.XiansuqiModel}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 限速器型号获取错误，已设置为默认值");
                Console.ResetColor();
                result.XiansuqiModel = "/";
            }
            
            try
            {
                Console.Write("请输入限速器编号: ");
                result.XiansuqiNum = Console.ReadLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 限速器编号: {result.XiansuqiNum}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 限速器编号获取错误，已设置为默认值");
                Console.ResetColor();
                result.XiansuqiNum = "/";
            }
            
            Console.WriteLine();
            Console.Write("请选择限速器方向 [0=单向, 1=双向]: ");
            string directionInput = Console.ReadLine();
            if (directionInput == "0")
            {
                result.XiansuqiDirection = "☑  单向 ☐  双向";
                result.xiansuqiDirectionForReport = "单向";
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("✓ 已选择: 单向");
            }
            else
            {
                result.XiansuqiDirection = "☐  单向 ☑  双向";
                result.xiansuqiDirectionForReport = "双向";
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("✓ 已选择: 双向");
            }
            Console.ResetColor();

            // ==================== 限速器铭牌速度参数录入 ====================
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("┌─────────────────────────────────────────┐");
            Console.WriteLine("│            限速器铭牌速度参数录入            │");
            Console.WriteLine("└─────────────────────────────────────────┘");
            Console.ResetColor();
            Console.WriteLine();

            // 电气动作速度
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("【电气动作速度】");
            Console.ResetColor();
            
            try
            {
                Console.Write("请输入限速器铭牌电气动作上行速度: ");
                result.XiansuqiElectricalUpSpeed = Console.ReadLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 电气动作上行速度: {result.XiansuqiElectricalUpSpeed}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 限速器铭牌电气动作上行速度获取错误，已设置为默认值");
                Console.ResetColor();
                result.XiansuqiElectricalUpSpeed = "/";
            }
            
            try
            {
                Console.Write("请输入限速器铭牌电气动作下行速度: ");
                result.XiansuqiElectricalDownSpeed = Console.ReadLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 电气动作下行速度: {result.XiansuqiElectricalDownSpeed}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 限速器铭牌电气动作下行速度获取错误，已设置为默认值");
                Console.ResetColor();
                result.XiansuqiElectricalDownSpeed = "/";
            }
            
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("【机械动作速度】");
            Console.ResetColor();
            
            try
            {
                Console.Write("请输入限速器铭牌机械动作上行速度: ");
                result.XiansuqiMechanicalUpSpeed = Console.ReadLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 机械动作上行速度: {result.XiansuqiMechanicalUpSpeed}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 限速器铭牌机械动作上行速度获取错误，已设置为默认值");
                Console.ResetColor();
                result.XiansuqiMechanicalUpSpeed = "/";
            }
            
            try
            {
                Console.Write("请输入限速器铭牌机械动作下行速度: ");
                result.XiansuqiMechanicalDownSpeed = Console.ReadLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"✓ 机械动作下行速度: {result.XiansuqiMechanicalDownSpeed}");
                Console.ResetColor();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("✗ 限速器铭牌机械动作下行速度获取错误，已设置为默认值");
                Console.ResetColor();
                result.XiansuqiMechanicalDownSpeed = "/";
            }

            // ==================== 数据录入完成 ====================
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("┌─────────────────────────────────────────┐");
            Console.WriteLine("│              数据录入完成                │");
            Console.WriteLine("└─────────────────────────────────────────┘");
            Console.ResetColor();
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("📊 正在生成Word报告，请稍候...");
            Console.ResetColor();
            Console.WriteLine();

            return result;
        }

        static void ObjsIndex(string str, JObject objs, out int indexj, out int indexi, out bool isContain)
        {

            indexi = 0;
            indexj = 0;
            isContain = false;

            var tableDetections = objs["Response"]["TableDetections"];

            var result = tableDetections
                .Select((table, j) => new { Table = table, J = j })
                .SelectMany(x => x.Table["Cells"]
                    .Select((cell, i) => new { Cell = cell, I = i, J = x.J }))
                .FirstOrDefault(x =>
                {
                    string cellText = x.Cell["Text"].ToString();
                    // 使用正则表达式进行精确匹配
                    return Regex.IsMatch(cellText, @"\b" + Regex.Escape(str) + @"\b");
                });

            if (result != null)
            {
                indexi = result.I;
                indexj = result.J;
                isContain = true;
            }
        }
    }
}