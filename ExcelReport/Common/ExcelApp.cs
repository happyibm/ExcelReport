using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelReport.Common
{
    /// <summary>
    /// Excel 处理类
    /// </summary>
    public class ExcelApp
    {
        public string mFilename;
        public Application app;
        public Workbooks wbs;
        public Workbook wb;
        public Worksheets wss;
        public Worksheet ws;                
        
        /// <summary>
        /// 打开一个Excel文件 
        /// </summary>
        /// <param name="FileName"></param>
        public void Open(string FileName)
        {
            app = new Application();
            wbs = app.Workbooks;
            wb = wbs.Open(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            mFilename = FileName;

            app.ScreenUpdating = false;
            app.EnableEvents = false;
            app.DisplayAlerts = false;
        }

        /// <summary>
        /// 获取一个工作表
        /// </summary>
        /// <param name="SheetName"></param>
        /// <returns></returns>
 
        public Worksheet GetSheet(string SheetName)
        {
            Worksheet s = (Worksheet)wb.Worksheets[SheetName];

            return s;
        }

        /// <summary>
        /// 添加一个工作表
        /// </summary>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public Worksheet AddSheet(string SheetName)
        {
            Worksheet s = (Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            s.Name = SheetName;
            s.PageSetup.HeaderMargin = app.InchesToPoints(0.196850393700787);\
            return s;
        }

        /// <summary>
        /// 删除一个工作表
        /// </summary>
        /// <param name="SheetName"></param> 
        public void DelSheet(string SheetName)
        {
            ((Worksheet)wb.Worksheets[SheetName]).Delete();
        }        

        /// <summary>
        /// 获取数组范围数据
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="begin"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public Array GetValues(Worksheet ws, string begin, string end)
        {
            var rang = ws.get_Range(begin, end);
            return (Array)rang.Cells.Value2;
        }

        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="ws"></param>
        public void RowDelete(Worksheet ws)
        {
            var count = ws.UsedRange.Rows.Count;
            var deleteRows = ws.get_Range("A1:A" + count.ToString(), Type.Missing).EntireRow;
            deleteRows.Delete(XlDeleteShiftDirection.xlShiftUp);
        }

        /// <summary>
        /// 拷贝行
        /// </summary>
        /// <param name="ws1"></param>
        /// <param name="ws2"></param>
        /// <param name="begin"></param>
        public void RowCopy(ref Worksheet ws1, ref Worksheet ws2, int rows, int begin)
        {
            var temp = ws2.get_Range("A1:A" + rows.ToString(), Type.Missing).EntireRow;
            var data = ws1.Rows.get_Item(begin, Type.Missing);

            temp.Copy(data);
        }
        
        /// <summary>
        /// 设值的工作表 X行Y列 value 值
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="value"></param>
        public void SetCellValue(ref Worksheet ws, int x, int y, object value)
        {
            ((Range)ws.Cells[x, y]).Value2 = value;
        }

        /// <summary>
        /// 设值的工作表 X行Y列 宽度  
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="value"></param>
        public void SetCellWidth(ref Worksheet ws, int x, int y, double value)
        {
            var colum = ((Range)ws.Cells[x, y]);
            colum.EntireColumn.ColumnWidth = value;
        }

        /// <summary>
        /// 保存文档 
        /// </summary>
        /// <returns></returns>
        public bool Save()
        {
            if (mFilename == " ")
            {
                return false;
            }
            else
            {
                try
                {
                    wb.Save();
                    return true;
                }

                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// 关闭一个Excel对象 销毁对象 
        /// </summary>
        public void Close()
        {
            app.ScreenUpdating = true;
            app.EnableEvents = true;

            app.DisplayAlerts = false;//提示是否保存

            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();

            //保存后退出，并释放资源
            ExcelKill.Kill(app);

            wb = null;
            wbs = null;
            app = null;
            wss = null;
            ws = null;
            GC.Collect();
        }
    }
}
