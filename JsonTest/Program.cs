using Newtonsoft.Json; //https://www.nuget.org/packages/Newtonsoft.Json
using Newtonsoft.Json.Linq;


namespace JsonTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            
            string json = File.ReadAllText(@"E:\\PythonProject\\限速器自动化\\识别结果\\2.json");
            JObject? objs = JObject.Parse(json);
            string deviceCode;
            deviceCode = objs["tables_result"][0]["body"][1]["words"].ToString();
            Console.WriteLine(deviceCode);
            bool a;
            a = objs["tables_result"][0]["body"][1]["words"].ToString().Contains(deviceCode);
            Console.WriteLine(a);
            
            
            var b = objs["tables_result"][0]["body"][1]["words"].ToString().Contains("日期");

            Console.WriteLine(objs["tables_result"][0]["body"].Count());

            int length = 0;

            foreach (var item in objs["tables_result"][0]["body"])
            {
                length++;
            }

            Console.WriteLine(length);

            for (var i = 0; i < length; i++)
            {
                var isContain = objs["tables_result"][0]["body"][i]["words"].ToString().Contains("设备代码");
                if (isContain)
                {
                    Console.WriteLine(i);
                    Console.WriteLine(objs["tables_result"][0]["body"][i+1]["words"].ToString());
                    Console.WriteLine("------------------------------------------------------");
                    return;
                }
            }



        }
    }
}
