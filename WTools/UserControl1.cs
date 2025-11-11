using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace WTools
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }
        DataTable dt = new DataTable();
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim().Length > 0)
            {
                dt.Rows.Clear();
                if (dataGridView1.Rows.Count > 0) dataGridView1.Rows.Clear();
                string parmat = "";
                string[] items = textBox1.Text.Split(',');
                for (int i = 0; i < items.Length; i++)
                {
                    if (i == 0)
                    {
                        parmat += "'" + items[i].Trim() + "'";
                    }
                    else
                    {
                        parmat += ",'" + items[i].Trim() + "'";
                    }
                }
                SqlConnection conn = new SqlConnection("Server =192.168.1.252;Database=WP01;User ID=sa;Password=dsc@53290529;encrypt=false;");
                SqlCommand cmd = new SqlCommand("SELECT TB002 製令,TB003 品號,[MB002] 品名,[MB004] 單位,[MB003] 規格,[TB004] 需求量,[MB064] 庫存量,[MB064]-[TB004] 差異 ,[MB092] 包裝數量 ,[TB009] 庫別  FROM [MOCTB] a  inner join [INVMB] b on a.TB003=b.MB001 where TB002 in (" + parmat + ") order by TB002", conn);
                cmd.Connection.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                dt.Load(rdr);
                cmd.Connection.Close();
                dataGridView1.DataSource = dt;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewRow row = dataGridView1.Rows[i];
                    if (Convert.ToInt32(row.Cells[7].Value) < 0)
                    {
                        row.Cells[7].Style.ForeColor = Color.Red;
                    }
                }
            }
            else
            {
                dt.Rows.Clear();
                if (dataGridView1.Rows.Count > 0) dataGridView1.Rows.Clear();
                dataGridView1.DataSource = dt;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel Files (*.xlsx)|";
                saveFileDialog1.DefaultExt = "xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");//指明非商业应用
                    using (var excel = new ExcelPackage())
                    {
                        // 建立分頁
                        var ws = excel.Workbook.Worksheets.Add("MySheet");
                        // 寫入資料試試
                        ws.Cells["A1"].Value = "製令";
                        ws.Cells["B1"].Value = "品號";
                        ws.Cells["C1"].Value = "品名";
                        ws.Cells["D1"].Value = "單位";
                        ws.Cells["E1"].Value = "規格";
                        ws.Cells["F1"].Value = "需求量";
                        ws.Cells["G1"].Value = "庫存量";
                        ws.Cells["H1"].Value = "差異";
                        ws.Cells["I1"].Value = "包裝數量";
                        ws.Cells["J1"].Value = "庫別";
                        ws.Cells["A2"].LoadFromDataTable(dt);
                        // 儲存 Excel
                        var file = new FileInfo(saveFileDialog1.FileName); // 檔案路徑
                        excel.SaveAs(file);
                        MessageBox.Show("轉檔完成.....");
                    }
                }
            }
            else
            {
                MessageBox.Show("查無製令庫存!!!!!");
            }
        }

    }
}
