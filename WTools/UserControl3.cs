using NiceLabel.SDK;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace WTools
{
    public partial class UserControl3 : UserControl
    {
        DataTable DT;
        IPrintEngine printEngine = PrintEngineFactory.PrintEngine;
        public UserControl3()
        {
            InitializeComponent();
        }

        private void UserControl3_Load(object sender, EventArgs e)
        {
            string filePath = "config.txt"; // 請替換為你的檔案實際路徑
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
                            textBox3.Text = tmp[1];
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

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            DT = new DataTable();
            string sql = "SELECT [TA025] 母製令,[TA002] 製令單號,[TA006] 品號,[TA034] 品名,[TA046] 單位";
            sql += ",case [TA011] when 'y' then '指定完工' when 'Y' then '已完工' when '1' then '未生產' when '2' then '已發料' end 狀態";
            sql += ",[TA015] 預計產量,[TA016] 已領套數,[TA017] 已生產量 FROM [MOCTA] ";
            sql += "where TA025 in (SELECT [TA025] FROM[WP01].[dbo].[MOCTA]";
            SqlConnection conn = new SqlConnection("Server =192.168.1.252;Database=WP01;User ID=sa;Password=dsc@53290529;encrypt=false;");

            if (textBox1.Text.Trim().Length == 11 && textBox2.Text.Trim().Length == 11)
            {
                sql += " where [TA002] BETWEEN '" + textBox1.Text.Trim() + "' AND '" + textBox2.Text.Trim() + "') order by [TA025],[TA002]";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Connection.Open();
                SqlDataReader sdr = cmd.ExecuteReader();
                if (sdr.HasRows)
                {
                    DT.Load(sdr);
                    dataGridView1.DataSource = DT;
                }
            }
            else
            {
                if (textBox1.Text.Trim().Length == 11)
                {
                    sql += " where [TA002] = '" + textBox1.Text.Trim() + "') order by [TA025],[TA002]";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Connection.Open();
                    SqlDataReader sdr = cmd.ExecuteReader();
                    if (sdr.HasRows)
                    {
                        DT.Load(sdr);
                        dataGridView1.DataSource = DT;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string uplable = "", quty = "0", ptname = "";
            string[] list = new string[4];
            if (textBox3.Text.Trim() == "")
            {
                MessageBox.Show("設定印表機名稱及複印數量!!!");
                return;
            }
            if (DT != null && DT.Rows.Count > 0 && DT.Rows.Count < 4)
            {
                int sno = 0;
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    DataRow dr = DT.Rows[i];
                    if (dr[0].ToString().Trim() != dr[1].ToString().Trim())
                    {
                        list[sno] = dr[1].ToString().Trim();
                        sno++;
                    }
                    else
                    {
                        uplable = dr[0].ToString().Trim();
                        quty = dr[6].ToString().Trim();
                        ptname = dr[3].ToString().Trim();
                    }
                }
                try
                {
                    string currentWorkingDirectory = Directory.GetCurrentDirectory() + "\\ppp.nlbl";
                    ILabel label = printEngine.OpenLabel(currentWorkingDirectory);

                    label.Variables["Pname"].SetValue(ptname);
                    label.Variables["Quty"].SetValue(quty);
                    label.Variables["Cmd1"].SetValue(uplable);
                    label.Variables["Cmd2"].SetValue(list[0]);
                    label.Variables["Cmd3"].SetValue(list[1]);
                    label.Variables["Cmd4"].SetValue(list[2]);
                    label.Variables["Cmd5"].SetValue(list[3]);

                    label.PrintSettings.PrinterName = textBox3.Text.Trim();
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
            else
            {
                MessageBox.Show("資料異常或數量過多!!!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = "config.txt";
            string contentToWrite = "Printer:" + textBox3.Text.Trim() + "\r\n" + "Count:" + numericUpDown1.Value.ToString() + "\r\n";
            try
            {
                File.WriteAllText(filePath, contentToWrite);
            }
            catch (Exception ex)
            {
                Console.WriteLine("設定失敗：" + ex.Message);
            }
        }
    }
}
