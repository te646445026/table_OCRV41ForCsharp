using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace table_OCRV41ForCsharp_net;

public class TencentOcrParser:IOcrParser
{
    public OcrResult Parse(string json)
    {
        var objs = JObject.Parse(json);
        OcrResult result = new OcrResult();

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


            Console.WriteLine("当前图片是: " + result.JianyanOrjiance);
        }
        catch
        {
            Console.WriteLine("获取检验还是检测失败,默认设置为检测");

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
        
        //string temperature;
        string jianyanOrjianceTiaojian;
        try
        {
            if (result.JianyanOrjiance.Equals("检验"))
            {
                jianyanOrjianceTiaojian = "检验条件";
            }
            else
            {
                jianyanOrjianceTiaojian = "检测条件";
            }
            int indexj;
            int indexi;
            bool isContain;
            ObjsIndex(jianyanOrjianceTiaojian, objs, out indexj, out indexi, out isContain);
            result.Temperature = objs["Response"]["TableDetections"][indexj]["Cells"][indexi + 1]["Text"].ToString().Replace("\n", "").Replace("\r", "");
            string temperature_pattern = @"\d{2,3}";
            MatchCollection temperatureNeed = Regex.Matches(result.Temperature, temperature_pattern);
            result.Temperature = $"温度：{temperatureNeed[0].ToString()}℃，  湿度：{temperatureNeed[1].ToString()}％ ， 电压：{temperatureNeed[2].ToString()}V";
            Console.WriteLine("温度、湿度、电压: " + result.Temperature);
        }
        catch
        {
            Console.WriteLine("检验或检测条件获取错误");
            result.Temperature = "温度：  ℃，湿度：  %，电压：  V";
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
            if (dateNeed != null)
            {
                result.Date = dateNeed[dateNeed.Count() - 1].Value;
                string date_or_month_pattern2 = @"\d+";
                MatchCollection matches = Regex.Matches(result.Date, date_or_month_pattern2);
                int year = int.Parse(matches[0].Value);
                int month = int.Parse(matches[1].Value);
                int day = int.Parse(matches[2].Value);
                result.Date = matches[0].Value + "年" + matches[1].Value + "月" + matches[2].Value + "日";
                Console.WriteLine("检验时间为：" + result.Date);
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
                result.NextYear = next_year_date.ToString("yyyy年MM月dd日");
                result.NextYearFlag = "";
                //计算审核校准日期
                DateTime shenhe_dateforcell = dateforcell.AddDays(1);
                result.ShenheDate = shenhe_dateforcell.ToString("yyyy年MM月dd日");
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

        //string xiansuqiModel;
        try
        {
            Console.WriteLine("输入限速器型号：");
            result.XiansuqiModel = Console.ReadLine();
            Console.WriteLine("限速器型号：" + result.XiansuqiModel);
        }
        catch
        {
            Console.WriteLine("限速器型号获取错误");
            result.XiansuqiModel = "/";
        }
        //string xiansuqiNum;
        try
        {
            Console.WriteLine("输入限速器编号：");
            result.XiansuqiNum = Console.ReadLine();
            Console.WriteLine("限速器编号：" + result.XiansuqiNum);
        }
        catch
        {
            Console.WriteLine("限速器编号获取错误");
            result.XiansuqiNum = "/";
        }
        //string xiansuqiDirection;
        //string xiansuqiDirectionForReport;
        Console.WriteLine("输入单向还是双向，0为单向，1为双向");
        if (Console.ReadLine() == "0")
        {
            result.XiansuqiDirection = "☑  单向 ☐  双向";
            result.xiansuqiDirectionForReport = "单向";
        }
        else
        {
            result.XiansuqiDirection = "☐  单向 ☑  双向";
            result.xiansuqiDirectionForReport = "双向";
        }

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