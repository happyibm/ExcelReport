using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ExcelReport
{
    /// <summary>
    /// 报表生成
    /// </summary>
    public partial class frmMain : Form
    {
        private event EventHandler StopThread;
        Thread thread;

        delegate void SetprogressBar(int max, int value, string info);
        SetprogressBar toSetprogressBar;

        string newFileName;

        public frmMain()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 选择excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
           var openDialog = new OpenFileDialog();
           openDialog.Filter = "excel 文件|*.xls;*.xlsx";

           if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
           {
               button1.Enabled = false;

               var fileName = openDialog.FileName;
               newFileName = fileName.Substring(0, fileName.LastIndexOf(".")) + "-打印版" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

               //文件拷贝
               System.IO.File.Copy(fileName, newFileName, true);

               toSetprogressBar = new SetprogressBar(SetprogressBarValue);

               thread = new Thread(new ThreadStart(DoWork));
               thread.IsBackground = true;
               thread.Start();

               StopThread += Form1_StopThread;

           }
        }

        /// <summary>
        /// 处理逻辑
        /// </summary>
        private void DoWork()
        {
            var list = new List<Model.Outstock>();

            var app = new Common.ExcelApp();
            app.Open(newFileName);

            //处理的sheet
            var dataSheet = app.GetSheet("数据");
            var tempSheet = app.GetSheet("模板");
            var reportSheet = app.AddSheet("报表");

            //获取数据
            var rowCount = dataSheet.UsedRange.Rows.Count;

            var values = app.GetValues(dataSheet, "A1", "I" + rowCount.ToString());

            Model.Outstock outstock;

            var datacount = values.GetLength(0) - 1;
            var maxcount = values.GetLength(0) * 2;


            this.progressBar1.Invoke(toSetprogressBar, new object[] { maxcount, 0, "开始读取数据..." });

            //出库数据
            for (int i = 2; i <= values.GetLength(0); i++)
            {
                outstock = new Model.Outstock();
                if (values.GetValue(i, 1).GetType().Name == "Double")
                {
                    outstock.RQ = DateTime.FromOADate(double.Parse(values.GetValue(i, 1).ToString())).ToString();
                }
                else
                {
                    outstock.RQ = values.GetValue(i, 1).ToString();
                }
                outstock.BH = values.GetValue(i, 2).ToString();
                outstock.MC = values.GetValue(i, 3).ToString();
                outstock.XH = values.GetValue(i, 4) == null ? "" : values.GetValue(i, 4).ToString();
                outstock.SL = values.GetValue(i, 5).ToString();
                outstock.DJ = values.GetValue(i, 6) == null ? "" : values.GetValue(i, 6).ToString();
                outstock.JE = values.GetValue(i, 7) == null ? "" : values.GetValue(i, 7).ToString();
                outstock.BGY = values.GetValue(i, 8) == null ? "" : values.GetValue(i, 8).ToString();
                outstock.JSR = values.GetValue(i, 9) == null ? "" : values.GetValue(i, 9).ToString();

                list.Add(outstock);

                this.progressBar1.Invoke(toSetprogressBar, new object[] { maxcount, i - 1, "共 " + datacount + " 条数据，已读取 " + (i - 1).ToString() + " 条" });
            }

            //根据模板生成数据
            int reportBegin = 1; //粘贴位置
            int reportRowBody = 8;//表体行数
            int reportHead = 3;//从表头行数
            int reportEnd = 6;//表尾行数
            string flagBH = string.Empty;//发票编号
            int flagIndex = 1;//表体索引
            int flagStep = 1;//表数据填充记录
            int flagCopy = 1;

            string rq = "     {0} 年  {1} 月  {2} 日           对方科目                     ";
            string bgr = "   主管                会计                保管员 {0}             经手人 {1} ";

            app.SetCellWidth(ref reportSheet, 1, 10, 11.25);
            app.SetCellWidth(ref reportSheet, 1, 7, 5.75);

            int flagP = 0;
            int flagC = list.Count();

            //循环处理生成报表
            foreach (var item in list)
            {
                //编号不一致属于另一张发票，拷贝模板
                if (string.Compare(item.BH, flagBH) != 0)
                {
                    flagBH = item.BH;

                    //拷贝模板
                    app.RowCopy(ref reportSheet, ref tempSheet, reportHead + reportRowBody + reportEnd, reportBegin);

                    //编号
                    app.SetCellValue(ref reportSheet, reportBegin, 8, "NO：" + item.BH);
                    //日期
                    var date = DateTime.Parse(item.RQ);
                    app.SetCellValue(ref reportSheet, reportBegin + 1, 4, string.Format(rq, date.Year, date.Month, date.Day));
                    //保管人
                    app.SetCellValue(ref reportSheet, reportBegin + 11, 1, string.Format(bgr, item.BGY, item.JSR));

                    //表体索引起点
                    flagIndex = reportBegin + reportHead;

                    //下一拷贝点
                    reportBegin += reportHead + reportRowBody + reportEnd;

                    flagCopy++;

                    flagStep = 1;
                }

                if (flagStep <= reportRowBody)
                {
                    //数据填充                       
                    app.SetCellValue(ref reportSheet, flagIndex, 1, item.MC);
                    app.SetCellValue(ref reportSheet, flagIndex, 3, item.XH);
                    app.SetCellValue(ref reportSheet, flagIndex, 7, item.SL);
                    //app.SetCellValue(ref reportSheet, flagIndex, 8, item.DJ);
                    //app.SetCellValue(ref reportSheet, flagIndex, 9, item.JE);
                }
                flagIndex++;

                if (flagStep == reportRowBody)
                {
                    //超过 reportRowBody 下一个模板填充
                    flagBH = string.Empty;
                    flagStep = 1;
                }

                flagStep++;

                flagP++;

                this.progressBar1.Invoke(toSetprogressBar, new object[] { maxcount, flagC + flagP, "共 " + datacount + " 条数据，已生成报表 " + flagP.ToString() + " 条" });
            }

            //删除数据表格
            app.DelSheet("数据");
            app.DelSheet("模板");

            app.Save();
            app.Close();

            this.progressBar1.Invoke(toSetprogressBar, new object[] { maxcount, maxcount, "共处理 " + datacount + " 条数据，报表生成完成。" });
        }

        /// <summary>
        /// 显示处理信息
        /// </summary>
        /// <param name="max"></param>
        /// <param name="value"></param>
        /// <param name="info"></param>
        private void SetprogressBarValue(int max,int value,string info)
        {
            this.progressBar1.Maximum = max;
            this.progressBar1.Value = value;
            this.label1.Text = info;

            if (this.progressBar1.Value == this.progressBar1.Maximum)
            {
                StopThread(this, new EventArgs());
            }
        }

        /// <summary>
        /// 停止线程
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Form1_StopThread(object sender, EventArgs e) 
        {             
            thread.Abort(); 
            MessageBox.Show("报表处理完成！");
            button1.Enabled = true;
        }
    }
}
