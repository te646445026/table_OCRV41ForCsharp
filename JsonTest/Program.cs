using Newtonsoft.Json; //https://www.nuget.org/packages/Newtonsoft.Json
using Newtonsoft.Json.Linq;


namespace JsonTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            KEY myKey = new KEY();

            string path = System.Environment.CurrentDirectory + @"\key.json";

            

            string key = File.ReadAllText(path);

            Console.WriteLine(key);

            myKey = JsonConvert.DeserializeObject<KEY>(key);

            Console.WriteLine(myKey.API_KEY);
            Console.WriteLine(myKey.SECRET_KEY);

        }
    }

    public class KEY
    {
        public string API_KEY { get; set; }
        public string SECRET_KEY { get; set; }

        private string KeyTOJson()
        {
            return JsonConvert.SerializeObject(this);
        }

        public void WriteKey(string path)
        {
            File.WriteAllText(path,KeyTOJson());
        }
    }
}
