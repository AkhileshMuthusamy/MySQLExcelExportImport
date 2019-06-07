using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Threading;
using System.Text.RegularExpressions;

namespace MySQL2Excel
{
    public partial class MainForm : Form
    {
        MySQLCon mysqlCon = new MySQLCon();
        string fileName;

        Exception exception = null;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.CheckFileExists = false;
                saveFileDialog.CheckPathExists = true;
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                saveFileDialog.AddExtension = true;
                saveFileDialog.FileName = "Export_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                if (saveFileDialog.ShowDialog() != DialogResult.Cancel)
                {
                    fileName = saveFileDialog.FileName;

                    btnGenerate.Enabled = false;
                    this.Cursor = Cursors.WaitCursor;
                    btnGenerate.Text = "Please Wait...";

                    ThreadStart threadStart = exportMySQLToExcel;
                    threadStart += () =>
                    {
                        // Executes when thread gets completed.
                        toolStripStatusLabel1.Text = "Completed";
                        if (exception is null)
                        {
                            this.Invoke(new Action(() => { MessageBox.Show(this, "Exported Data to excel!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); }));
                        }
                        else
                        {
                            this.Invoke(new Action(() => { MessageBox.Show(this, "Failed to export data." + Environment.NewLine + exception, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }));
                        }

                        this.Invoke(new Action(() => { this.Cursor = Cursors.Default; }));
                        btnGenerate.Invoke(new Action(() => {
                            btnGenerate.Enabled = true;
                            btnGenerate.Text = "Export to Excel";
                        }));
                    };

                    Thread t = new Thread(threadStart);
                    t.IsBackground = true;
                    t.Start();
                    
                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.ToString());
                MessageBox.Show(ex.ToString());
            }
        }

        void exportMySQLToExcel()
        {
            MySqlConnection conn = null;
            MySqlDataReader rdr = null;

            try
            {
                conn = new MySqlConnection(mysqlCon.connectionString);
                conn.Open();
                
                using (ExcelPackage objExcelPackage = new ExcelPackage())
                {
                    Int32 exlCol = 1;
                    ExcelWorksheet objWorksheet = objExcelPackage.Workbook.Worksheets.Add("Data");
                    objWorksheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 10));

                    string stm = "SHOW COLUMNS FROM uio_250;";
                    MySqlCommand cmd = new MySqlCommand(stm, conn);
                    rdr = cmd.ExecuteReader();

                    while (rdr.Read())
                    {                          
                        objWorksheet.Cells[1, exlCol].Value = rdr.IsDBNull(0) ? null : rdr.GetString(0);
                        exlCol++;
                    }

                    rdr.Close();

                    stm = "SELECT * FROM uio_250;";
                    cmd = new MySqlCommand(stm, conn);
                    rdr = cmd.ExecuteReader();

                    Int32 columnCount = rdr.FieldCount;
                    Int32 row = 2;

                    while (rdr.Read())
                    {
                        for (Int32 col = 0; col < columnCount; col++)
                        {
                            exlCol = col + 1;

                            objWorksheet.Cells[row, exlCol].Value = rdr.IsDBNull(col) ? null : rdr.GetString(col);
                        }
                        Console.WriteLine(row.ToString());
                        this.toolStripStatusLabel1.GetCurrentParent().Invoke(new Action(() => toolStripStatusLabel1.Text = "Rows Processed: " + row.ToString()));
                        row++;
                    }

                    rdr.Close();
                    rdr.Dispose();


                    try
                    {
                        FileStream objFileStrm = File.Create(fileName);
                        objFileStrm.Close();

                        this.toolStripStatusLabel1.GetCurrentParent().Invoke(new Action(() => toolStripStatusLabel1.Text = "Saving " + row.ToString() + " records in excel file."));
                        //Write content to excel file    
                        File.WriteAllBytes(fileName, objExcelPackage.GetAsByteArray());

                        
                    }
                    catch (Exception ex)
                    {
                        exception = ex;
                    }

                }

            }
            catch (Exception ex)
            {
                exception = ex;
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }

            }
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            
            if(folderBrowserDialog.ShowDialog() != DialogResult.Cancel)
            {
                string appointmentfolderPath = $"{folderBrowserDialog.SelectedPath}\\File Type\\appointment";
                string[] dirFiles = Directory.GetFiles(appointmentfolderPath);
                foreach (string exfile in dirFiles)
                {
                    string dataFile = exfile;
                    if (exfile.Substring(dataFile.Length - 1) == "s")
                    {
                        ConvertToXlsx(dataFile);
                        dataFile = dataFile + "x";
                    }

                    importExcelToMySQL(dataFile);
                    
                }
                MessageBox.Show("Import process completed");
            }

        }

        void importExcelToMySQL(string excelFile)
        {
            MySqlConnection conn = null;
            string executeQuery = null;

            try
            {
                byte[] bin = File.ReadAllBytes(excelFile);

                conn = new MySqlConnection(mysqlCon.connectionString);
                conn.Open();

                //create a new Excel package in a memorystream
                using (MemoryStream stream = new MemoryStream(bin))
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    //loop all worksheets
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        //loop all rows
                        for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                        {
                            executeQuery = @"
INSERT INTO appointment
(
`branch`,
`appointment_id`,
`appointment_d_t`,
`stall`,
`service_start_d_t`,
`service_end_d_t`,
`service_type`,
`service_item`,
`car_wash`,
`car_in_out_type`,
`car_in_d_t`,
`car_out_d_t`,
`remarks`,
`customer_code`,
`customer_name`,
`tel_no`,
`mobile_no`,
`postal_code`,
`address`,
`model`,
`reg_no`,
`vin`,
`mileage`,
`service_adviser`,
`sales_staff`,
`file_type_excel`
)
VALUES
(
";
                            //loop all columns in a row
                            for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                            {
                                //add the cell data to the List
                                if (worksheet.Cells[i, j].Value != null)
                                {
                                    string trimValue = worksheet.Cells[i, j].Value.ToString();
                                    trimValue = Regex.Replace(trimValue, @"[^\w\d\s\.\,\:\]\[\#\)\(\+\=\*\@\-\/]", " ");
                                    if (worksheet.Cells[i, j].Value.ToString().Length >= 255)
                                    {                                                                               
                                        executeQuery = executeQuery + "'" + (trimValue.Substring(0, 255)) + "',";
                                    }
                                    else
                                    {
                                        executeQuery = executeQuery + "'" + trimValue + "',";
                                    }
                                }
                                else
                                {
                                    executeQuery = executeQuery + "NULL,";
                                }
                            }
                            executeQuery = executeQuery + "'appointment',";
                            executeQuery = executeQuery.Substring(0, executeQuery.Length - 1);
                            executeQuery = executeQuery + ");";
                            //MessageBox.Show(executeQuery);

                            MySqlCommand myCommand = new MySqlCommand(executeQuery, conn);
                            myCommand.ExecuteNonQuery();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                exception = ex;
                MessageBox.Show(exception.ToString());
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }

            }        
        }

        public void ConvertToXlsx(String file)
        { 
            var app = new Microsoft.Office.Interop.Excel.Application();
            var wb = app.Workbooks.Open(file);
            wb.SaveAs(Filename: file + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();     
        }
    }
}
