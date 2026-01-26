using System.Collections.Generic;
using System.Linq;

namespace table_OCRV41ForCsharp_net_framework.Models
{
    public class PlaceholderMapping
    {
        public string Placeholder { get; set; }
        public string PropertyName { get; set; }
        public bool AppendUnit { get; set; }
        public string Unit { get; set; }
        public bool NeedsPrefix { get; set; }
        public string Prefix检验 { get; set; }
        public string Prefix检测 { get; set; }
    }

    public static class PlaceholderMappings
    {
        public static readonly List<PlaceholderMapping> Mappings = new List<PlaceholderMapping>
        {
            new PlaceholderMapping { Placeholder = "[1]", PropertyName = "ReportNum", NeedsPrefix = true, Prefix检验 = "D", Prefix检测 = "E" },
            new PlaceholderMapping { Placeholder = "[2]", PropertyName = "UserName" },
            new PlaceholderMapping { Placeholder = "[3]", PropertyName = "Date" },
            new PlaceholderMapping { Placeholder = "[4]", PropertyName = "UserName" },
            new PlaceholderMapping { Placeholder = "[5]", PropertyName = "MaintenanceUnit" },
            new PlaceholderMapping { Placeholder = "[6]", PropertyName = "UsingAddress" },
            new PlaceholderMapping { Placeholder = "[7]", PropertyName = "ElevatorDeviceType" },
            new PlaceholderMapping { Placeholder = "[8]", PropertyName = "DeviceCode" },
            new PlaceholderMapping { Placeholder = "[9]", PropertyName = "SerialNum" },
            new PlaceholderMapping { Placeholder = "[10]", PropertyName = "Speed", AppendUnit = true, Unit = "m/s" },
            new PlaceholderMapping { Placeholder = "[11]", PropertyName = "XiansuqiManufacturingUnit" },
            new PlaceholderMapping { Placeholder = "[12]", PropertyName = "XiansuqiModel" },
            new PlaceholderMapping { Placeholder = "[13]", PropertyName = "XiansuqiNum" },
            new PlaceholderMapping { Placeholder = "[14]", PropertyName = "XiansuqiDirection" },
            new PlaceholderMapping { Placeholder = "[15]", PropertyName = "XiansuqiElectricalUpSpeed", AppendUnit = true, Unit = "m/s" },
            new PlaceholderMapping { Placeholder = "[16]", PropertyName = "XiansuqiElectricalDownSpeed", AppendUnit = true, Unit = "m/s" },
            new PlaceholderMapping { Placeholder = "[17]", PropertyName = "XiansuqiMechanicalUpSpeed", AppendUnit = true, Unit = "m/s" },
            new PlaceholderMapping { Placeholder = "[18]", PropertyName = "XiansuqiMechanicalDownSpeed", AppendUnit = true, Unit = "m/s" },
            new PlaceholderMapping { Placeholder = "[19]", PropertyName = "ShenheDate" },
            new PlaceholderMapping { Placeholder = "[20]", PropertyName = "NextYearFlag" },
        };
    }
}
