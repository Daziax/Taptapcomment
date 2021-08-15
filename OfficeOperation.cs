using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.IO;
namespace taptapcomment
{
    public class ExcelOp
    {
        public void AddRaws(ExcelPackage excel, ref int usedRowsCount, int usedColsCount, ref string[] data)
        {
            ExcelWorksheet ws = excel.Workbook.Worksheets[0];
            //Workbook wb,Worksheet ws;
            try
            {

                for (int i = 0; i < usedColsCount; ++i)//
                {
                    ws.Cells[usedRowsCount + 1, i + 1].Value = data[i];
                }
                ++usedRowsCount;
            }
            catch (Exception ex)
            {
                excel.Save();
                Console.WriteLine("添加列时出错:"+ex.Message);
                //CloseProcess(filePath, excel, wb);
            }
        }

        public void CreateExcelFile(string filePath, string sheetName, string[] cols)
        {
            //create
            object Nothing = System.Reflection.Missing.Value;
            var fileStream = new FileStream(filePath, FileMode.Create);
            var app=new ExcelPackage(fileStream);
            ExcelWorkbook workbook = app.Workbook;
            ExcelWorksheet worksheet = workbook.Worksheets.Add(sheetName);

            worksheet.Name = sheetName;

            try
            {

                //headline
                for (int i = 0; i < cols.Length; ++i)
                {
                    worksheet.Cells[1, i + 1].Value = cols[i];
                }

                //FileInfo fileInfo=new FileInfo (filePath);

                app.SaveAs(fileStream);
                //workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception e)
            {
                Console.WriteLine("创建文件异常:" + e);
            }
            finally
            {
                fileStream.Dispose();
                app.Dispose();
            }
            //CloseProcess(filePath, app, workbook);
        }


        #region  .Framework版本

        /*public void AddRaws( Application excel,Workbook wb,Worksheet ws,int usedRowsCount,ref string[] data)
        {
            try
            {
                usedRowsCount = ws.UsedRange.Rows.Count;//赋值有效行
                      
                for (int i = 0; i <= ws.UsedRange.Columns.Count; ++i)//
                {
                    if (ws.Rows[usedRowsCount] != null)
                    {
                        ws.Cells[usedRowsCount+1,i+1]=data[i];
                    }
                }
            }
            catch (Exception ex) 
            { 
                Console.WriteLine(ex.Message, "error");
                //CloseProcess(filePath, excel, wb);
            }
        }
        public object[] Init(string filePath)
        {
            object[] result=new object[4];
            Application excel = new Application();
            excel.Visible = false;
            Workbook wb = excel.Workbooks.Open(filePath);
            Worksheet ws = (Worksheet)excel.Worksheets[1]; //索引从1开始 //(Excel.Worksheet)wb.Worksheets["SheetName"];
            int usedRowsCount = ws.UsedRange.Rows.Count;//有效行，索引从1开始
            result[0]=(object)excel;
            result[1]=(object)wb;
            result[2]=(object)ws;
            result[3]=(object)usedRowsCount;
            return result;
        }
        public void CreateExcelFile(string filePath, string sheetName, string[] cols)
        {
            //create
            object Nothing = System.Reflection.Missing.Value;
            var app = new Application();
            app.Visible = false;
            Workbook workbook = app.Workbooks.Add(Nothing);
            Worksheet worksheet = (Worksheet)workbook.Sheets[1];

            worksheet.Name = sheetName;

            //headline
            for (int i = 0; i < cols.Length; ++i)
            {
                worksheet.Cells[1, i + 1] = cols[i];
            }

            //worksheet.Columns[2].NumberFormatLocal = "@";//设置第二列为 文本格式
            workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            CloseProcess(filePath, app, workbook);
        }
        /// <summary>
        /// 关闭Excel进程
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="excel"></param>
        /// <param name="workbook"></param>
        public void CloseProcess(string excelPath, Application excel, Workbook workbook)
        {
            Process[] localByNameApp = Process.GetProcessesByName(excelPath);//获取程序名的所有进程
            if (localByNameApp.Length > 0)
            {
                foreach (var app in localByNameApp)
                {
                    if (!app.HasExited)
                    {
                        #region
                        ////设置禁止弹出保存和覆盖的询问提示框   
                        excel.DisplayAlerts = false;
                        excel.AlertBeforeOverwriting = false;

                        ////保存工作簿   
                        excel.Application.Workbooks.Add(true).Save();
                        ////保存excel文件   
                        excel.Save();
                        ////确保Excel进程关闭   
                        excel.Quit();
                        excel = null;
                        #endregion
                        app.Kill();//关闭进程  
                    }
                }
            }
            if (workbook != null)
                workbook.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            // 安全回收进程
            System.GC.GetGeneration(excel);
        }
    */
        #endregion
    }
}