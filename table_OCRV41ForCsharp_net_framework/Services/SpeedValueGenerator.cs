using System;
using System.Collections.Generic;
using System.Linq;
using table_OCRV41ForCsharp_net_framework.Models;

namespace table_OCRV41ForCsharp_net_framework.Services
{
    public interface ISpeedValueGenerator
    {
        void GenerateTestValues(OcrResult result);
    }

    public class SpeedValueGenerator : ISpeedValueGenerator
    {
        private readonly Random _random;

        /// <summary>
        /// 额定速度 → 各组取值范围（电气上/下，机械上/下）
        /// 数据来源于《电梯限速器动作速度取值表》
        /// </summary>
        private static readonly Dictionary<double, SpeedRangeSet> ReferenceTable =
            new Dictionary<double, SpeedRangeSet>
            {
                [0.50] = new SpeedRangeSet
                {
                    ElectricalUp = (0.65, 0.71),
                    ElectricalDown = (0.64, 0.70),
                    MechanicalUp = (0.72, 0.78),
                    MechanicalDown = (0.73, 0.79)
                },
                [0.63] = new SpeedRangeSet
                {
                    ElectricalUp = (0.74, 0.79),
                    ElectricalDown = (0.73, 0.78),
                    MechanicalUp = (0.75, 0.80),
                    MechanicalDown = (0.75, 0.80)
                },
                [1.00] = new SpeedRangeSet
                {
                    ElectricalUp = (1.21, 1.30),
                    ElectricalDown = (1.20, 1.29),
                    MechanicalUp = (1.31, 1.40),
                    MechanicalDown = (1.32, 1.41)
                },
                [1.50] = new SpeedRangeSet
                {
                    ElectricalUp = (1.82, 1.90),
                    ElectricalDown = (1.81, 1.89),
                    MechanicalUp = (1.92, 1.98),
                    MechanicalDown = (1.93, 1.99)
                },
                [1.60] = new SpeedRangeSet
                {
                    ElectricalUp = (1.95, 2.04),
                    ElectricalDown = (1.90, 1.99),
                    MechanicalUp = (2.01, 2.10),
                    MechanicalDown = (2.09, 2.18)
                },
                [1.75] = new SpeedRangeSet
                {
                    ElectricalUp = (2.09, 2.18),
                    ElectricalDown = (2.08, 2.17),
                    MechanicalUp = (2.21, 2.30),
                    MechanicalDown = (2.22, 2.31)
                },
                [2.00] = new SpeedRangeSet
                {
                    ElectricalUp = (2.39, 2.48),
                    ElectricalDown = (2.38, 2.47),
                    MechanicalUp = (2.51, 2.60),
                    MechanicalDown = (2.52, 2.61)
                },
                [2.50] = new SpeedRangeSet
                {
                    ElectricalUp = (2.88, 3.08),
                    ElectricalDown = (2.85, 3.05),
                    MechanicalUp = (3.08, 3.18),
                    MechanicalDown = (3.11, 3.21)
                },
            };

        public SpeedValueGenerator()
        {
            _random = new Random();
        }

        public void GenerateTestValues(OcrResult result)
        {
            double ratedSpeed = ParseRatedSpeed(result.Speed);
            SpeedRangeSet ranges = LookupReferenceRange(ratedSpeed);
            bool isSingleDirection = (result.XiansuqiDirection ?? "").Contains("单向");

            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("┌─────────────────────────────────────────┐");
            Console.WriteLine("│            实测速度值自动生成              │");
            Console.WriteLine("└─────────────────────────────────────────┘");
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine($"额定速度: {ratedSpeed:F2} m/s");
            Console.WriteLine("规则: 电气动作速度 < 机械动作速度，上下行有偏移");

            // 1. 电气上行
            GenerateSpeedGroup("电气上行", ranges.ElectricalUp,
                result.XiansuqiElectricalUpSpeed,
                v => SetTestValues(v, result,
                    (r, s) => r.ElectricalUpTest1 = s,
                    (r, s) => r.ElectricalUpTest2 = s,
                    (r, s) => r.ElectricalUpTest3 = s),
                avg => result.ElectricalUpAvg = avg);

            // 2. 电气下行
            GenerateSpeedGroup("电气下行", ranges.ElectricalDown,
                result.XiansuqiElectricalDownSpeed,
                v => SetTestValues(v, result,
                    (r, s) => r.ElectricalDownTest1 = s,
                    (r, s) => r.ElectricalDownTest2 = s,
                    (r, s) => r.ElectricalDownTest3 = s),
                avg => result.ElectricalDownAvg = avg);

            // 3. 机械上行（单向限速器跳过）
            if (isSingleDirection)
            {
                Console.WriteLine($"  机械上行: 单向限速器，跳过");
                result.MechanicalUpTest1 = "/";
                result.MechanicalUpTest2 = "/";
                result.MechanicalUpTest3 = "/";
                result.MechanicalUpAvg = "/";
            }
            else
            {
                GenerateSpeedGroup("机械上行", ranges.MechanicalUp,
                    result.XiansuqiMechanicalUpSpeed,
                    v => SetTestValues(v, result,
                        (r, s) => r.MechanicalUpTest1 = s,
                        (r, s) => r.MechanicalUpTest2 = s,
                        (r, s) => r.MechanicalUpTest3 = s),
                    avg => result.MechanicalUpAvg = avg);
            }

            // 4. 机械下行
            GenerateSpeedGroup("机械下行", ranges.MechanicalDown,
                result.XiansuqiMechanicalDownSpeed,
                v => SetTestValues(v, result,
                    (r, s) => r.MechanicalDownTest1 = s,
                    (r, s) => r.MechanicalDownTest2 = s,
                    (r, s) => r.MechanicalDownTest3 = s),
                avg => result.MechanicalDownAvg = avg);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("✓ 实测速度值生成完毕");
            Console.ResetColor();
        }

        /// <summary>
        /// 生成一组速度的3次实测值，并计算平均值
        /// 公式：单次 = 下限 + (上限 - 下限) * RAND()
        /// </summary>
        private void GenerateSpeedGroup(
            string groupName,
            (double Lower, double Upper) range,
            string nameplateSpeedStr,
            Action<(double, double, double)> setTestValues,
            Action<string> setAvg)
        {
            double lower = range.Lower;
            double upper = range.Upper;

            // 如果铭牌值有效且在范围内，优先以铭牌值为参考生成
            double effectiveLower = lower;
            double effectiveUpper = upper;
            if (!string.IsNullOrEmpty(nameplateSpeedStr) && nameplateSpeedStr != "/"
                && double.TryParse(nameplateSpeedStr, out double nameplate))
            {
                if (nameplate >= lower && nameplate <= upper)
                {
                    // 铭牌值在范围内，以铭牌值为中心，保留±0.03偏移空间
                    effectiveLower = Math.Max(lower, nameplate - 0.03);
                    effectiveUpper = Math.Min(upper, nameplate + 0.03);
                    Console.WriteLine($"  {groupName}: 铭牌值 {nameplate:F2} 在范围内 [{lower:F2}, {upper:F2}]，以其为中心生成");
                }
                else
                {
                    Console.WriteLine($"  {groupName}: 铭牌值 {nameplate:F2} 超出范围 [{lower:F2}, {upper:F2}]，使用范围随机");
                }
            }

            // 生成3次测量值：下限 + (上限 - 下限) * RAND()
            double v1 = GenerateSingleValue(effectiveLower, effectiveUpper);
            double v2 = GenerateSingleValue(effectiveLower, effectiveUpper);
            double v3 = GenerateSingleValue(effectiveLower, effectiveUpper);

            // 排序，使显示更有序（可选）
            double[] sorted = new[] { v1, v2, v3 }.OrderBy(x => x).ToArray();
            v1 = sorted[0];
            v2 = sorted[1];
            v3 = sorted[2];

            setTestValues((v1, v2, v3));

            // 平均值保留2位小数
            double avg = Math.Round((v1 + v2 + v3) / 3.0, 2);
            setAvg(avg.ToString("F2"));

            Console.WriteLine($"    {v1:F2} / {v2:F2} / {v3:F2} → 平均 {avg:F2} m/s (范围 {lower:F2}~{upper:F2})");
        }

        /// <summary>
        /// 生成单个测量值：下限 + (上限 - 下限) * RAND()
        /// </summary>
        private double GenerateSingleValue(double lower, double upper)
        {
            double value = lower + (upper - lower) * _random.NextDouble();
            return Math.Round(value, 2);
        }

        /// <summary>
        /// 设置3次测试值
        /// </summary>
        private void SetTestValues(
            (double v1, double v2, double v3) values,
            OcrResult result,
            Action<OcrResult, string> set1,
            Action<OcrResult, string> set2,
            Action<OcrResult, string> set3)
        {
            set1(result, values.v1.ToString("F2"));
            set2(result, values.v2.ToString("F2"));
            set3(result, values.v3.ToString("F2"));
        }

        /// <summary>
        /// 解析额定速度，从字符串中提取数值
        /// </summary>
        private double ParseRatedSpeed(string speedStr)
        {
            if (string.IsNullOrEmpty(speedStr) || speedStr == "/")
                return 1.00;

            if (double.TryParse(speedStr, out double result))
                return result;

            var match = System.Text.RegularExpressions.Regex.Match(speedStr, @"(\d+\.?\d*)");
            if (match.Success && double.TryParse(match.Groups[1].Value, out double parsed))
                return parsed;

            return 1.00;
        }

        /// <summary>
        /// 根据额定速度查找参照表
        /// 精确匹配 > 最接近的较大值 > 按比例推算
        /// </summary>
        private SpeedRangeSet LookupReferenceRange(double ratedSpeed)
        {
            // 精确匹配
            if (ReferenceTable.TryGetValue(ratedSpeed, out var exact))
                return exact;

            // 最接近的较大值
            var closestUp = ReferenceTable.Keys
                .Where(k => k >= ratedSpeed)
                .OrderBy(k => k)
                .FirstOrDefault();

            if (closestUp > 0)
            {
                Console.WriteLine($"    额定速度 {ratedSpeed:F2} 无精确匹配，使用 {closestUp:F2} 的参照范围");
                return ReferenceTable[closestUp];
            }

            // 超出最大范围，按比例推算
            double maxRated = ReferenceTable.Keys.Max();
            double ratio = ratedSpeed / maxRated;
            var maxRange = ReferenceTable[maxRated];
            Console.WriteLine($"    额定速度 {ratedSpeed:F2} 超出参照表，按比例推算");
            return new SpeedRangeSet
            {
                ElectricalUp = ScaleRange(maxRange.ElectricalUp, ratio),
                ElectricalDown = ScaleRange(maxRange.ElectricalDown, ratio),
                MechanicalUp = ScaleRange(maxRange.MechanicalUp, ratio),
                MechanicalDown = ScaleRange(maxRange.MechanicalDown, ratio)
            };
        }

        private (double Lower, double Upper) ScaleRange((double Lower, double Upper) range, double ratio)
        {
            return (Math.Round(range.Lower * ratio, 2), Math.Round(range.Upper * ratio, 2));
        }

        /// <summary>
        /// 各组速度范围集合
        /// </summary>
        private class SpeedRangeSet
        {
            public (double Lower, double Upper) ElectricalUp { get; set; }
            public (double Lower, double Upper) ElectricalDown { get; set; }
            public (double Lower, double Upper) MechanicalUp { get; set; }
            public (double Lower, double Upper) MechanicalDown { get; set; }
        }
    }
}
