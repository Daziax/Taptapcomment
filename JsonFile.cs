using System.Diagnostics;
using System.IO;
using System;
using System.Text;
namespace taptapcomment
{
    public class JsonFile
    {
        /// <summary>
        /// 创建文件
        /// </summary>
        /// <param name="filePath"></param>
        internal void Create(string filePath)
        {

            if (!File.Exists(filePath))
            {
                FileStream fileStream = File.Create(filePath);
                fileStream.Dispose();
            }
            else
            {
                Console.WriteLine("该路径已存在文件，是否覆盖？");
                if (Console.ReadLine() == "y")
                {
                    FileStream fileStream = File.Create(filePath);
                    fileStream.Dispose();
                    Console.WriteLine("已经覆盖原文件");
                }
                else
                    Console.WriteLine("创建失败");
            }
        }
        internal void Write(string filePath, string content)
        {

            using (FileStream fileStream = File.OpenWrite(filePath))
            {
                //StreamReader streamReader=new StreamReader (fileStream);
                byte[] buffer = new UTF8Encoding().GetBytes(content);
                fileStream.Write(buffer, 0, buffer.Length);
            }
            Console.WriteLine("成功将内容写入文件。");

        }
    }
}
