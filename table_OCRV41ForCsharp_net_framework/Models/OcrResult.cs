namespace table_OCRV41ForCsharp_net_framework.Models
{
    public class OcrResult
    {
        public string Text { get; set; }
        public string DeviceCode { get; set; }
        public string Model { get; set; }
        public string SerialNum { get; set; }
        public string ManufacturingUnit { get; set; }
        public string UserName { get; set; }
        public string UsingAddress { get; set; }
        public string MaintenanceUnit { get; set; }
        public string Speed { get; set; }
        public string RatedLoad { get; set; }

        public string ReportNum { get; set; }
        public string Date { get; set; }
        public string NextYear { get; set; }
        public string XiansuqiModel { get; set; }
        public string XiansuqiNum { get; set; }
        public string XiansuqiDirection { get; set; }
        public string xiansuqiDirectionForReport { get; set; }
        public string NextYearFlag { get; set; }
        public string ShenheDate { get; set; }
        public string JianyanOrjiance { get; set; }
        public string XiansuqiManufacturingUnit { get; set; }
        public string XiansuqiElectricalUpSpeed { get; set; }
        public string XiansuqiElectricalDownSpeed { get; set; }
        public string XiansuqiMechanicalUpSpeed { get; set; }
        public string XiansuqiMechanicalDownSpeed { get; set; }
        public string ElevatorDeviceType { get; set; }

        // 实测速度值 — 平均值 [21]-[24]
        public string ElectricalUpAvg { get; set; }
        public string ElectricalDownAvg { get; set; }
        public string MechanicalUpAvg { get; set; }
        public string MechanicalDownAvg { get; set; }

        // 实测速度值 — 第1/2/3次 [25]-[36]
        public string ElectricalUpTest1 { get; set; }
        public string ElectricalUpTest2 { get; set; }
        public string ElectricalUpTest3 { get; set; }
        public string ElectricalDownTest1 { get; set; }
        public string ElectricalDownTest2 { get; set; }
        public string ElectricalDownTest3 { get; set; }
        public string MechanicalUpTest1 { get; set; }
        public string MechanicalUpTest2 { get; set; }
        public string MechanicalUpTest3 { get; set; }
        public string MechanicalDownTest1 { get; set; }
        public string MechanicalDownTest2 { get; set; }
        public string MechanicalDownTest3 { get; set; }
    }
}