using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Renci.SshNet;
using System.Net.Sockets;
using System.Diagnostics;
using System.Threading;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace Ping監控IP_background
{
    public partial class Form1 : Form
    {
        //==========基礎定義區==========
        public Form1()
        {
            InitializeComponent();
            // 背景程序
            backgroundWorker1.WorkerReportsProgress = true;

        }

        // 建立 Excel程序 物件
        Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        // 建立excel 物件
        Excel.Workbook xlWorkBook;
        // 建立excel工作頁 物件
        Excel.Worksheet xlWorkSheet;
        // 建立 OLT 字典 物件
        Dic OLT_DIC = new Dic();

        object misValue = System.Reflection.Missing.Value;

        // 變數
        int RowCount; //高
        string OLT; //OLT_ID
        string OLT_IP; //OLT IP
        string ONT_IP; //ONT IP
        string result; //Ping結果

        int i; // 執行變數
        string line; //輸出變數
        int ping_true =0;
        int ping_false=0;

        //==============================



        //==========按鈕==========
        private void Button1_Click(object sender, EventArgs e)
        {

            if (textBox2.Text=="")
            {
                MessageBox.Show("還沒選擇檔案呢! 急什麼!");
            }
            else
            {
                //開關按鈕
                button1.Enabled = false;
                button2.Enabled = true;
                textBox1.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox2.Enabled = false;
                button3.Enabled = false;


                // 匯入OLT字典
                OLT_DIC.CreateDictionary();
                label6.Text = "執行輸出：";
                richTextBox1.Text = "執行開始，請稍等。";

                //指定excel物件為哪份excel位置
                xlWorkBook = xlApp.Workbooks.Open(textBox2.Text);

                //讀取該份excel工作頁
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(int.Parse(textBox1.Text));

                // 抓高
                RowCount = xlWorkSheet.UsedRange.Rows.Count;

                //執行背景程序
                backgroundWorker1.RunWorkerAsync();
            }
            

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.ShowDialog();
            textBox2.Text = file.FileName;

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            ////xlWorkBook.Close(true, misValue, misValue);
            ////xlApp.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);
            backgroundWorker1.CancelAsync();
            button1.Enabled = true;
            button2.Enabled = false;
            
        }

        //==============================



        //==========backgroundwork==========
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            for (i = 2; i < RowCount + 1; i++)
            {
                // 取OLT 值
                OLT = ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[i, textBox3.Text]).Text.ToString();
                // 抓字典當中OLT IP
                OLT_IP = (OLT_DIC.read(OLT));

                using (var client = new SshClient(OLT_IP, 22, "admin", "123"))
                {
                    // 建立連線

                    try
                    {
                    client.Connect();
                    var stream = client.CreateShellStream("", 0, 0, 0, 0, 0);
                    // Send the command
                    Thread.Sleep(4500);
                    stream.WriteLine("en");
                    Thread.Sleep(100);
                    stream.WriteLine("ping");
                    stream.WriteLine("");
                    ONT_IP = ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[i, textBox4.Text]).Text.ToString();
                    stream.WriteLine(ONT_IP);
                    stream.WriteLine("");
                    stream.WriteLine("1");
                    stream.WriteLine("64");
                    stream.WriteLine("2");
                    stream.WriteLine("");
                        while ((line = stream.ReadLine(TimeSpan.FromSeconds(6))) != null) //如果輸出的值不是null則持續輸出。
                        {
                            //Console.WriteLine(line);

                            Thread.Sleep(300);
                            //IndexOF 會抓後面關鍵字，在line 中第幾個位置
                            int id = line.IndexOf("100% packet loss");
                            int id2 = line.IndexOf("1 received");

                            // 所以>0就代表有抓到，不等於-1就代表抓到了
                            if (id != -1)
                            {
                                result = "執行結果：沒通!";
                                xlWorkSheet.Cells[i, textBox5.Text] = "沒通";
                                ping_false++;
                            }

                            if (id2 != -1)
                            {
                                result = "執行結果：有通!";
                                xlWorkSheet.Cells[i, textBox5.Text] = "有通";
                                ping_true++;
                            }

                        }
                    }
                    catch (System.Runtime.InteropServices.InvalidComObjectException)
                    {

                      
                    }
                    
                    
                    //stream.WriteLine("echo 'sample command output'");

                    // Read with a suitable timeout to avoid hanging




                }
                backgroundWorker1.ReportProgress(i);
            }
            Thread.Sleep(5000);
            xlWorkBook.Save();
            xlWorkBook.Close(true, misValue, misValue);

            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);


            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("完成!");
            button1.Enabled = true;
            textBox1.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
            textBox2.Enabled = true;
            button3.Enabled = true;

        }



        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            richTextBox1.AppendText("\r\n" + "======================");
            richTextBox1.AppendText("\r\n" + "目前載入的OLT為" + OLT);
            richTextBox1.AppendText("\r\n" + "目前載入的OLT IP為" + OLT_IP);
            richTextBox1.AppendText("\r\n" + "Ping的IP為：" + ONT_IP);
            richTextBox1.AppendText("\r\n" + result);
            richTextBox1.AppendText("\r\n" + "======================");
            

        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            richTextBox1.AppendText("\r\n" + "執行完成!");
            richTextBox1.AppendText("\r\n" + "Ping成功數量："+ping_true);
            richTextBox1.AppendText("\r\n" + "Ping失敗數量："+ping_false);
            button1.Enabled = true;
            button1.Enabled = false;
            textBox1.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
            textBox2.Enabled = true;
            button3.Enabled = true;
        }







        //=============================

    }
}
