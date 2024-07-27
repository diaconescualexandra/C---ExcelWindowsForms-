using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;




namespace WindowsFormsApp1
{

    public partial class CitireExcel : Form
    {
        //private SpreadsheetControl spreadsheetControl;
        public DataTable currentTable = new DataTable();
        private ExcelData excelData;



        
           


        public CitireExcel()
        {
            InitializeComponent();

            excelData = new ExcelData();

        }



        public void ReadExcel(string filePath)
        {
            try
            {
                // Load a workbook from a stream.
                using (FileStream stream = new FileStream(filePath, FileMode.Open))
                {
                    spreadsheetControl1.LoadDocument(stream, DocumentFormat.Xlsx);
                }



                currentTable.Clear();
                currentTable.Columns.Clear();
                IWorkbook workbook = spreadsheetControl1.Document;
                Worksheet workSheet = workbook.Worksheets[0];
                CellRange usedRange = workSheet.GetUsedRange();

                for (int i = 0; i < usedRange.ColumnCount; i++)
                {
                    currentTable.Columns.Add("Column" + i.ToString());
                }

                for (int i = 0; i < usedRange.RowCount; i++)
                {
                    DataRow newRow = currentTable.NewRow();
                    for (int j = 0; j < usedRange.ColumnCount; j++)
                    {
                        newRow[j] = workSheet.Cells[i, j].DisplayText;
                    }
                    currentTable.Rows.Add(newRow);
                }



                foreach (DataRow row in currentTable.Rows)
                {
                    int ctEmptyColPerRows = excelData.emptyColumnsPerRow(row);
                    int ctFullColPerRows = excelData.fullColumnsPerRow(row);
                    int totalNoCol = excelData.totalNumberOfColumns(currentTable);
                    int totalNoRows = excelData.totalNumberOfRows(currentTable);

                }
                excelData.dt = currentTable;
            }

            catch (Exception ex)
            {
                excelData.SetException(ex.Message);
                string message = excelData.CatchAllExceptions();
                Console.WriteLine(message);
            }

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //InitializeSpreadsheet(xtraOpenFileDialog1.FileName);
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (xtraOpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                barEditItem3.EditValue = xtraOpenFileDialog1.FileName;
                excelData.SetFilePath(xtraOpenFileDialog1.FileName);
                string filePath = excelData.GetFilePath();
                string extensionName = excelData.WhichExtension(filePath);
                ReadExcel(xtraOpenFileDialog1.FileName);
            }
        }

        private void barEditItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            int b = 2;
        }

        public DataTable returnDataTable()
        {
            
            foreach (DataRow row in currentTable.Rows)
            {
                foreach (DataColumn col in currentTable.Columns)
                {
                    Console.WriteLine($"{col.ColumnName}: {row[col]}");
                }
            }

            return currentTable;
        }
    }



}




