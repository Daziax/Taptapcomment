using System.Net.WebSockets;
using System.Net.Http;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Threading;
using System.Net.Security;
using System;
using System.Net;
using System.IO;

namespace taptapcomment
{
    public class Client
    {
        public static HttpClient client = new HttpClient();
      
        private Client(){}
        

    }
    public class HttpHandler
    {
        HttpClient client = Client.client;
        internal async void GetComment()
        {
            client.Timeout = TimeSpan.FromSeconds(5);
            /*client.DefaultRequestHeaders.Host = "www.jianshu.com";
            
            client.DefaultRequestHeaders.Referrer = new Uri("https://www.jianshu.com/");
            client.DefaultRequestHeaders.Add("Accept", "application/json");
            client.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate, br");
            client.DefaultRequestHeaders.Add("Accept-Language", "");
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36");*/
            //int count =0;
            string rspText; //= await client.GetStringAsync("https://www.taptap.com/webapiv2/review/v2/by-app?app_id=170078&limit=10&from="+count+"&X-UA=V%3D1%26PN%3DWebApp%26LANG%3Dzh_CN%26VN_CODE%3D38%26VN%3D0.1.0%26LOC%3DCN%26PLT%3DPC%26DS%3DAndroid%26UID%3Dd83aeb12-e9a6-4277-81cb-daf9d8b8a327%26DT%3DPC");
            for(int count=0;count<20; count+=10)
            {
                rspText = await client.GetStringAsync("https://www.taptap.com/webapiv2/review/v2/by-app?app_id=170078&limit=10&from="+count+"&X-UA=V%3D1%26PN%3DWebApp%26LANG%3Dzh_CN%26VN_CODE%3D38%26VN%3D0.1.0%26LOC%3DCN%26PLT%3DPC%26DS%3DAndroid%26UID%3Dd83aeb12-e9a6-4277-81cb-daf9d8b8a327%26DT%3DPC");
                Console.WriteLine("=========================================");
                Console.WriteLine(rspText);
            }
            
        }
       
    }
}