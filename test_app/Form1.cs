using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;

namespace test_app
{
    public partial class Form1 : Form
    {
        private SqlConnection connection; 

        public Form1()
        {
            
            InitializeComponent();            
                      
        }

        private void ReadData()
        {
            try
            {
                string connectionString = "Data Source=DESKTOP-5MVT38T\\SQLEXPRESS;Initial Catalog=INTERMECH_BASE;Integrated Security=True";
                connection = new SqlConnection(connectionString);
                string query = "SELECT IMS_OBJECTS.F_OBJECT_ID, IMS_OBJECT_TYPES.F_OBJ_NAME, IMS_OBJECTS_VIEW.CAPTION FROM IMS_OBJECTS LEFT JOIN IMS_OBJECT_TYPES on IMS_OBJECTS.F_LC_STEP = IMS_OBJECT_TYPES.F_OBJECT_TYPE LEFT JOIN IMS_OBJECTS_VIEW on IMS_OBJECTS.F_ID =  IMS_OBJECTS_VIEW.F_ID;";
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();

                DataTable dataTable = new DataTable();
                dataTable.Load(reader);
                dataTable.Columns[0].ColumnName = "ID";
                dataTable.Columns[1].ColumnName = "Name Type";
                dataTable.Columns[2].ColumnName = "Caption";

                dataGridView1.DataSource = dataTable;
                
                reader.Close();
                dataGridView1.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
                
            }
            catch (SqlException ex)
            {
                Console.WriteLine("Ошибка подключения к базе данных: " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка: " + ex.Message);
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ReadData();

                }

      

        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            
            string searchValue = textBoxSearch.Text.ToLower();

            DataTable dataTable = (DataTable)dataGridView1.DataSource;
            if (dataTable != null)
            {
                string filter = $"CONVERT(ID, 'System.String') LIKE '%{searchValue}%' OR [Name Type] LIKE '%{searchValue}%' OR [Caption] LIKE '%{searchValue}%'";

                dataTable.DefaultView.RowFilter = filter;
            }
        }

        private void dataGridView1_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           Enabled = false;
            if (e.RowIndex >= 0 && e.ColumnIndex == -1)
            {
                try
                {
                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                        DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];
                        
                        string connectionString = "Data Source=DESKTOP-5MVT38T\\SQLEXPRESS;Initial Catalog=INTERMECH_BASE;Integrated Security=True";
                        connection = new SqlConnection(connectionString);
                        string query = "SELECT IMS_OBJECTS.F_OBJECT_ID as ИД, IMS_OBJECT_TYPES.F_OBJ_NAME as Наименование_Типа, IMS_OBJECTS_VIEW.CAPTION as Описание, IMS_RELATION_TYPES.F_DESCRIPTION as Отношение_к_родителю   FROM IMS_RELATIONS LEFT JOIN IMS_OBJECTS ON IMS_RELATIONS.F_PART_ID = IMS_OBJECTS.F_ID " +
                            "LEFT JOIN IMS_OBJECT_TYPES on IMS_OBJECTS.F_LC_STEP = IMS_OBJECT_TYPES.F_OBJECT_TYPE LEFT JOIN IMS_OBJECTS_VIEW on IMS_OBJECTS.F_OBJECT_ID = IMS_OBJECTS_VIEW.F_OBJECT_ID " +
                            "LEFT JOIN IMS_RELATION_TYPES on IMS_RELATIONS.F_RELATION_TYPE = IMS_RELATION_TYPES.F_RELATION_TYPE WHERE IMS_RELATIONS.F_PROJ_ID = '" + selectedRow.Cells[0].Value + "';";
                        SqlCommand command = new SqlCommand(query, connection);
                        connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        DataTable dataTable = new DataTable();
                        dataTable.Load(reader);
                        worksheet.Cells[1, 1, 1, 4].Merge = true;
                        worksheet.Cells[1, 1, 1, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells[1, 1].Value = "Состав объекта № "+ selectedRow.Cells[0].Value;
                        worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        for (int i = 0; i < dataTable.Columns.Count; i++)
                        {
                            worksheet.Cells[2, i + 1].Value = dataTable.Columns[i].ColumnName;
                            worksheet.Cells[2, i+1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells[2, i + 1].EntireColumn.Width= 30;
                        }
                        int rowIndex = 3;
                        foreach (DataRow row in dataTable.Rows)
                        {
                            for (int columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                            {
                                worksheet.Cells[rowIndex, columnIndex + 1].Value = row[columnIndex];
                                worksheet.Cells[rowIndex, columnIndex + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            }
                            rowIndex++;
                        }
                        reader.Close();

                     
                        var tempFilePath = Path.GetTempFileName() + ".xlsx";
                        excelPackage.SaveAs(new FileInfo(tempFilePath));
                        System.Diagnostics.Process.Start(tempFilePath);
                    }
                
                  
                   

                }
                catch (SqlException ex)
                {
                    Console.WriteLine("Ошибка подключения к базе данных: " + ex.Message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
                finally
                {
                    if (connection.State == ConnectionState.Open)
                    {
                        connection.Close();
                    }
                    Enabled = true;
                }

               
              
                   
                
            }
        }
    }
}

