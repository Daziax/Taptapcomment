using System.Net.Mime;
using System.Reflection.PortableExecutable;
using System.Text.RegularExpressions;
using System.Text;
using System.Net.Http;
using System;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;

namespace taptapcomment
{
    class Program
    {
        static void Main(string[] args)
        {
            HttpClient client = Client.client;
            //HttpHandler handler=new HttpHandler ();
            client.Timeout = TimeSpan.FromSeconds(5);
            string rspText = string.Empty; //= await client.GetStringAsync("https://www.taptap.com/webapiv2/review/v2/by-app?app_id=170078&limit=10&from="+count+"&X-UA=V%3D1%26PN%3DWebApp%26LANG%3Dzh_CN%26VN_CODE%3D38%26VN%3D0.1.0%26LOC%3DCN%26PLT%3DPC%26DS%3DAndroid%26UID%3Dd83aeb12-e9a6-4277-81cb-daf9d8b8a327%26DT%3DPC");
                                           //Stream stream;
            JsonFile jsonFile = new JsonFile();
            string filePath = "/users/dazai/project/files/taptapcomment.xlsx";//"/users/dazai/project/files/taptapcomment.json";
            Console.WriteLine("运行成功");
            //jsonFile.Create(filePath);
            
            //string delimited = @"\G(.+)[\t\u007c](.+)\r?\n";
            //Regex regexContent = new Regex("\"extended_entities[0-9a-zA-Z:\"\\\\<>&; \\[\\]\\{\\}\\n,_]+\"collapsed\"");
            Regex regexContent = new Regex("\"extended_entities(.*?)collapse");
            Regex regexText = new Regex("\"text\":\"(.*?)\"");//("\"text\":\"[0-9a-zA-Z\\\\&;<> ]*\"");
            Regex regexScore = new Regex("\"score\":[0-9],");
            Regex regexDevice = new Regex("\"device\":\"(.*?)\"");//("\"device\":\"[a-zA-Z0-9\\\\ ]*\"");
            Regex regexPlayedDuration = new Regex("\"played_tips\":\"(.*?)\"");//("\"played_tips\":\"[a-zA-Z0-9 \\\\]*\"");
            Regex regexCreatedTime = new Regex("\"created_time\":[0-9 ]*,");
            Regex regexUpdatedTime = new Regex("\"updated_time\":[0-9 ]*,");
            Regex regexIsExistent = new Regex("\"success\":[ ]*false");
            MatchCollection matchCollection;
            //StringBuilder result = new StringBuilder();
            int count = 0;
            bool isExistent = true;
           
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            string[] data = new string[6];
            ExcelOp excelOp = new ExcelOp();
            excelOp.CreateExcelFile(filePath, "Comment", new string[6] { "创建日期", "编辑日期", "游戏时长", "评分", "设备", "评论" });
            FileInfo fileInfo=new FileInfo (filePath);
            ExcelPackage excel = new ExcelPackage (fileInfo);
            //excel.Visible = false;
            ExcelWorkbook wb = excel.Workbook;
            ExcelWorksheet ws = wb.Worksheets[0];
            int usedRowsCount = ws.Dimension.Rows;//有效行，索引从1开始
            
            try
            {
                while (isExistent)
                {
                    try
                    {
                        rspText = client.GetStringAsync("https://www.taptap.com/webapiv2/review/v2/by-app?app_id=170078&limit=10&from=" + count + "&X-UA=V%3D1%26PN%3DWebApp%26LANG%3Dzh_CN%26VN_CODE%3D38%26VN%3D0.1.0%26LOC%3DCN%26PLT%3DPC%26DS%3DAndroid%26UID%3Dd83aeb12-e9a6-4277-81cb-daf9d8b8a327%26DT%3DPC").Result;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message + "\ncount=" + count);
                        break;
                    }

                    matchCollection = regexContent.Matches(rspText);

                    //匹配是否是最后一页
                    if (regexIsExistent.IsMatch(rspText))
                        isExistent = false;
                    ///提取每页内容
                    for (int i = 0; i < matchCollection.Count; i++)
                    {
                        //string a0=matchCollection[i].Value;
                        //string a1=regexCreatedTime.Match(a0).Value;
                        //string a2=Regex.Match(a1,"[0-9]+").Value;
                        data[0] = TransTime(Regex.Match(regexCreatedTime.Match(matchCollection[i].Value).Value, "[0-9]+").Value);
                        data[1] = TransTime(Regex.Match(regexUpdatedTime.Match(matchCollection[i].Value).Value, "[0-9]+").Value);
                        data[2] = Regex.Unescape(regexPlayedDuration.Match(matchCollection[i].Value).Value);
                        data[3] = Regex.Unescape(regexScore.Match(matchCollection[i].Value).Value);
                        data[4] = Regex.Unescape(regexDevice.Match(matchCollection[i].Value).Value);
                        data[5] = Regex.Unescape(regexText.Match(matchCollection[i].Value).Value);
                        excelOp.AddRaws(excel,ref usedRowsCount,6, ref data);
                    }
                    excel.Save();
                    count += 10;
                    if(count%100==0)
                         Console.WriteLine("count:"+count);
                    
                   
                    //break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("循环为单元格赋值出错:"+e.Message);
            }
            finally
            {
                excel.Dispose();
            }
           
            //jsonFile.Write(filePath, result.ToString());
            Console.WriteLine("爬取结束!");
        }
        private static string TransTime(string st)
        {
            DateTime nowTime;
            ulong str = Convert.ToUInt64(st);
            if (str.ToString().Length == 13)
            {
                nowTime = new DateTime(1970, 1, 1, 8, 0, 0).AddMilliseconds(str);
            }
            else
            {
                nowTime = new DateTime(1970, 1, 1, 8, 0, 0).AddSeconds(str);
            }
            return nowTime.ToString();
        }
    }
}
