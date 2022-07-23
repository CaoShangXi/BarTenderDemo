using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BarTenderDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string serialNumber = textBox1.Text;
            string origin = textBox2.Text;
            using (StreamWriter stream = new StreamWriter(path + "\\param.txt"))
            {
                stream.Write(serialNumber + "," + origin);
            }
            string filePath = path + "\\test.btw";
            PrintMethod1(filePath, "1");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string serialNumber = textBox1.Text;
            string origin = textBox2.Text;
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\test02.btw";
            PrintMethod2(filePath, "2", serialNumber, origin);
        }

        /// <summary>
        /// BarTender列印方法1
        /// </summary>
        /// <param name="destFilePath">文件路径</param>
        /// <param name="copies">列印份数</param>
        public void PrintMethod1(string destFilePath, string copies)
        {
            Process p = new Process();
            p.StartInfo.FileName = "bartend.exe";
            //列印btw檔案並最小化程序
            p.StartInfo.Arguments = $@"/AF={destFilePath} /P /min=SystemTray";
            p.EnableRaisingEvents = true;
            int pageCount = Convert.ToInt32(copies);
            for (int i = 0; i < pageCount; i++)
            {
                p.Start();
            }
        }

        /// <summary>
        /// BarTender列印方法2
        /// </summary>
        /// <param name="destFilePath">文件路径</param>
        /// <param name="copies">列印份数</param>
        /// <param name="serialNumber">序号</param>
        /// <param name="origin">产地</param>
        public void PrintMethod2(string destFilePath, string copies, string serialNumber, string origin)
        {
            BarTender.Application btApp = new BarTender.Application();
            BarTender.Format btFormat = btApp.Formats.Open(destFilePath, false, "");
            btFormat.NumberSerializedLabels = Convert.ToInt32(copies);//序列标签数
            btFormat.SetNamedSubStringValue("serialNumber", serialNumber);//条码
            btFormat.SetNamedSubStringValue("origin", origin);//文字
            btFormat.PrintOut(false, false);
            btFormat.Close(BarTender.BtSaveOptions.btSaveChanges);
        }
    }
}
