using System;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;
using table_OCRV41ForCsharp_net_framework.Interfaces;
using table_OCRV41ForCsharp_net_framework.Models;

namespace table_OCRV41ForCsharp_net_framework.Services
{
    // 定义 API 密钥服务
    public class KeyService : IKeyService
    {
        public KEY CheckKey()
        {
            KEY myKey = new KEY();
            string keyPath = System.Environment.CurrentDirectory + @"\key.json";
            if (!File.Exists(keyPath))
            {
                MessageBox.Show("密钥文件缺失,点击确认后手动输入");

                Console.WriteLine("请输入SecretId");
                do
                {
                    myKey.SecretId = Console.ReadLine();
                } while (myKey.SecretId == null);

                Console.WriteLine("请输入SecretKey");
                do
                {
                    myKey.SecretKey = Console.ReadLine();
                } while (myKey.SecretKey == null);

                string keyJson = JsonConvert.SerializeObject(myKey);

                File.WriteAllText(keyPath, keyJson);

            }
            else
            {
                string keyJson = File.ReadAllText(keyPath);
                myKey = JsonConvert.DeserializeObject<KEY>(keyJson);
            }

            return myKey;
        }
    }
}