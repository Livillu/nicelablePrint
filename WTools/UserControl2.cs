using NiceLabel.SDK;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WTools
{
    public partial class UserControl2 : UserControl
    {
        DataTable DT;
        IPrintEngine printEngine = PrintEngineFactory.PrintEngine;
        public UserControl2()
        {
            InitializeComponent();
        }

        private void GetPURTH()
        {
            dataGridView1.DataSource = null;
            DT = new DataTable();
            string whereSql = "";
            string sql = "SELECT [TH002] 單號,[TH004] 品號,[TH005] 品名,[TH006] 規格,[TH007] 數量,[TH008] 單位,[TH010] 批號,[TH014] 驗收日 FROM [WP01].[dbo].[PURTH]";
            if (TB001.Text.Trim().Length > 0)
            {
                whereSql += " And TH001='" + TB001.Text.Trim() + "'";
            }
            if (TB002_1.Text.Trim().Length > 0 && TB002_2.Text.Trim().Length > 0)
            {
                whereSql += " And ([TH002] between '" + TB002_1.Text.Trim() + "' and '" + TB002_2.Text.Trim() + "')";
            }
            if (checkBox1.Checked && TB019_1.Text.Trim().Length > 0 && TB019_2.Text.Trim().Length > 0)
            {
                whereSql += " And ([TH014] between '" + TB019_1.Text.Trim() + "' and '" + TB019_2.Text.Trim() + "')";
            }
            SqlConnection conn = new SqlConnection("Server =192.168.1.252;Database=WP01;User ID=sa;Password=dsc@53290529;encrypt=false;");
            if (whereSql.Length > 3)
            {
                sql += " WHERE TH004<>'999' " + whereSql;
            }
            else
            {
                sql += " WHERE 1<>1";
            }
                sql += whereSql;
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Connection.Open();
            SqlDataReader sdr = cmd.ExecuteReader();
            if (sdr.HasRows)
            {
                DT.Load(sdr);
                dataGridView1.DataSource = DT;
            }
            else
            {
                DT = null;
            }
        }

        private void GetINVTB()
        {
            dataGridView1.DataSource = null;
            DT = new DataTable();
            string whereSql = "";
            string sql = "SELECT [TB002] 單號,[TB004] 品號,[TB005] 品名,[TB006] 規格,[TB007] 數量,[TB008] 單位,[TB014] 批號,[TB015] 有效日期  FROM [WP01].[dbo].[INVTB] WHERE TB001 IN ('117','118','11A','11B') ";
            if (TB001.Text.Trim().Length > 0)
            {
                whereSql += " And TB001='" + TB001.Text.Trim() + "'";
            }
            if (TB002_1.Text.Trim().Length > 0 && TB002_2.Text.Trim().Length > 0)
            {
                whereSql += " And ([TB002] between '" + TB002_1.Text.Trim() + "' and '" + TB002_2.Text.Trim() + "')";
            }
            if (checkBox1.Checked && TB019_1.Text.Trim().Length > 0 && TB019_2.Text.Trim().Length > 0)
            {
                whereSql += " And ([CREATE_DATE] between '" + TB019_1.Text.Trim() + "' and '" + TB019_2.Text.Trim() + "')";
            }
            /*
            if (checkBox1.Checked && TB019_1.Text.Trim().Length > 0 && TB019_2.Text.Trim().Length > 0)
            {
                if (whereSql != "") whereSql += " And ";
                whereSql += "(TB019 between '"+ TB019_1.Text.Trim() + "' and '"+ TB019_2.Text.Trim() + "')";
            }
            */
            SqlConnection conn = new SqlConnection("Server =192.168.1.252;Database=WP01;User ID=sa;Password=dsc@53290529;encrypt=false;");

            sql += whereSql;
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Connection.Open();
            SqlDataReader sdr = cmd.ExecuteReader();
            if (sdr.HasRows)
            {
                DT.Load(sdr);
                dataGridView1.DataSource = DT;
            }
            else
            {
                DT = null;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            GetPURTH();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string Pdate = "", Pcode = "", Pname = "";
            if (textBox4.Text.Trim() == "")
            {
                MessageBox.Show("設定印表機名稱及複印數量!!!");
                return;
            }
            PrinterSettings printerSettings = new PrinterSettings();
            printerSettings.PrinterName = textBox4.Text.Trim();
            
            string currentWorkingDirectory = Directory.GetCurrentDirectory() + "\\ppp2.nlbl";
            ILabel label = printEngine.OpenLabel(currentWorkingDirectory);
            label.PrintSettings.PrinterName= textBox4.Text.Trim();
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(item.Cells[0].Value))
                {
                    try
                    {
                        Pcode = item.Cells[2].Value.ToString();
                        Pname = item.Cells[3].Value.ToString();
                        Pdate = item.Cells[7].Value.ToString();
                        label.Variables["Pcode"].SetValue(Pcode);
                        label.Variables["Pname"].SetValue(Pname);
                        label.Variables["Pdate"].SetValue(Pdate);

                        label.PrintSettings.PrinterName = textBox4.Text.Trim();
                        try
                        {
                            label.Print(Convert.ToInt16(numericUpDown1.Value));
                        }
                        catch
                        {
                            label.PrintSettings.Reset();
                            label.Print(Convert.ToInt16(numericUpDown1.Value));
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Exception: " + ex.Message);
                        return;
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = "config2.txt";
            string contentToWrite = "Printer:" + TB002_2.Text.Trim() + "\r\n" + "Count:" + numericUpDown1.Value.ToString() + "\r\n";
            try
            {
                File.WriteAllText(filePath, contentToWrite);
            }
            catch (Exception ex)
            {
                Console.WriteLine("設定失敗：" + ex.Message);
            }
        }

        private void UserControl2_Closed(object sender, FormClosedEventArgs e)
        {
            printEngine.Shutdown();
        }

        private void UserControl2_Load_1(object sender, EventArgs e)
        {
            string filePath = "config2.txt"; // 請替換為你的檔案實際路徑
            try
            {
                string content = File.ReadAllText(filePath);
                using (StringReader reader = new StringReader(content))
                {
                    string line;
                    // Read lines until the end of the string is reached (ReadLine returns null)
                    while ((line = reader.ReadLine()) != null)
                    {
                        var tmp = line.Split(':');
                        if (tmp[0] == "Printer")
                        {
                            textBox4.Text = tmp[1];
                        }
                        if (tmp[0] == "Count")
                        {
                            numericUpDown1.Value = Convert.ToInt16(tmp[1]);
                        }
                    }
                    printEngine.Initialize();
                }

            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("檔案未找到！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤： " + ex.Message);
            }
        }
    }
}
