using System;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;

namespace ExcelDemo
{
    public partial class Form1 : Form
    {
        System.Data.DataTable allRecords;
        public Form1()
        {
            InitializeComponent();
            insertExcel.Visible = false;
        }

        private void btnReadExcel_Click(object sender, EventArgs e)
        {
            try
            {
                XLWorkbook wb = new XLWorkbook();
                System.Data.DataTable dt = Read("select * from [DemoDB].[dbo].[Marksheets]");
                wb.Worksheets.Add(dt, "ExcelFiles");
                string folderPath = "D:\\Excel Files\\";
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                string path = Path.Combine("D:\\Excel Files\\", $"{DateTime.Now.Ticks.ToString("X")}_Excel.xlsx");
                wb.SaveAs(path);
                Display($"File is downloaded at following path: {path}");
            }
            catch (Exception exception)
            {
                Display(exception.Message.ToString());
            }
        }

        private void openExcel_Click(object sender, EventArgs e)
        {
            try
            {
                string fname = "";
                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Title = "Excel File Dialog";
                fdlg.InitialDirectory = @"d:\";
                fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    fname = fdlg.FileName;
                }
                allRecords = new System.Data.DataTable();
                string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
                using (XLWorkbook workBook = new XLWorkbook(fname))
                {
                    IXLWorksheet workSheet = workBook.Worksheet(1);
                    bool firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                allRecords.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            allRecords.Rows.Add();
                            int i = 0;
                            foreach (IXLCell cell in row.Cells())
                            {
                                allRecords.Rows[allRecords.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }
                        dataGridView1.DataSource = allRecords;
                        insertExcel.Visible = true;
                    }
                }
            }
            catch (Exception exception)
            {
                insertExcel.Visible = false;
                allRecords = null;
                Display(exception.Message.ToString());
            }
        }

        public void Display(string msg)
        {
            MessageBox.Show(msg);
        }

        public System.Data.DataTable Read(string query)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();
            string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dataTable);
            conn.Close();
            da.Dispose();
            return dataTable;
        }

        private void clearExcel_Click(object sender, EventArgs e)
        {
            allRecords = null;
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            insertExcel.Visible = false;
            Display("Records cleared Successfully");
        }

        private void insertExcel_Click(object sender, EventArgs e)
        {
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlBulkCopy bulk = new SqlBulkCopy(connection))
                    {
                        bulk.DestinationTableName = "[DemoDB].[dbo].[Marksheets]";
                        bulk.ColumnMappings.Add(0, 1);
                        bulk.ColumnMappings.Add(1, 2);
                        bulk.ColumnMappings.Add(2, 3);
                        bulk.ColumnMappings.Add(3, 4);
                        bulk.WriteToServer(allRecords);
                        Display("Data inserted successfully");
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        allRecords = null;
                        insertExcel.Visible = false;
                    }
                }
            }
            catch (Exception excep)
            {
                insertExcel.Visible = false;
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                allRecords = null;
                Display(excep.Message.ToString());
            }
        }
    }
}
