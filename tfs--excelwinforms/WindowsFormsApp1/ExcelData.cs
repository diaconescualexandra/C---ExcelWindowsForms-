using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class ExcelData
    {
        public  DataTable dt { get; set; }
        private string exception;
        private string filePath { get; set; }
        //private DataTable excelDataTable { get; set; }
        private int contorEmptyPerRow { get; set; }
        private int contorFullPerRow { get; set; }
        private string extensionName { get; set; }



        public ExcelData()
        {
            contorEmptyPerRow = 0;
            contorFullPerRow = 0;
            exception = "";
            filePath = "";
            extensionName = "";

        }

        public int emptyColumnsPerRow(DataRow row)
        {
            contorEmptyPerRow = 0;
            foreach (string _valoare in row.ItemArray)
            {
                if (String.IsNullOrEmpty(_valoare))
                {
                    contorEmptyPerRow += 1;
                }
            }
            return contorEmptyPerRow;

        }

        public int fullColumnsPerRow(DataRow row)
        {
            contorFullPerRow = 0;
            foreach (string _valoare2 in row.ItemArray)
            {
                if (!String.IsNullOrEmpty(_valoare2))
                {
                    contorFullPerRow += 1;
                }

            }
            return contorFullPerRow;

        }
        public int totalNumberOfRows(DataTable db)
        {
            return db.Rows.Count;
        }

        public int totalNumberOfColumns(DataTable db)
        {
            return db.Columns.Count;
        }

        public void SetFilePath(string path)
        {
            filePath = path;

        }

        public string GetFilePath()
        {
            return filePath;
        }

        public void SetException(string ex)
        {
            exception = ex;
        }
        public string CatchAllExceptions()
        {
            return exception;
        }
        public string WhichExtension(string filePath)
        {
            if (isXLS(filePath) == true)
            {
                extensionName = ".xls";
            }
            else extensionName = ".xlsx";

            return extensionName;

        }

        public bool isXLS(string filePath)
        {
            if (Path.GetExtension(filePath) == ".xls")
            {
                return true;
            }
            else return false;
        }

        public string readFile(string filePath)
        {
            string readContents;
            using (StreamReader streamReader = new StreamReader(filePath, Encoding.UTF8))
            {
                readContents = streamReader.ReadToEnd();
            }
            return readContents;
        }
        //public DataTable returnDataTable()
        //{
        //    //try
        //    //{
        //        foreach (DataRow row in dt.Rows)
        //        {
        //            foreach (DataColumn col in dt.Columns)
        //            {
        //                Console.WriteLine($"{col.ColumnName}: {row[col]}");
        //            }
        //        }
        //    //}catch (Exception ex) { return ""; }

        //    return dt;
        //}

    }
}
