using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace ConsoleApplication1
{

    class Program
    {
        static void ConvertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {
            string cnnStr = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"", excelFilePath);
            OleDbConnection cnn = new OleDbConnection(cnnStr);
            DataTable dt = new DataTable();
            cnn.Open();
            var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
            string sql = String.Format("select * from [{0}]", worksheet);
            var da = new OleDbDataAdapter(sql, cnn);
            da.Fill(dt);
            cnn.Close();
            
            using (var wtr = new StreamWriter(csvOutputFile)) 
            {
                foreach (DataRow row in dt.Rows) 
                {
                    bool firstLine = true;
                    foreach (DataColumn col in dt.Columns) 
                    {
                        if (!firstLine) 
                            wtr.Write(","); 
                        else
                            firstLine = false; 
                        var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                        wtr.Write(String.Format("\"{0}\"", data));
                    }
                    wtr.WriteLine();
                } 
            }
        }

        static void Main(string[] args)
        {

            ConvertExcelToCsv("C:\\Users\\Developer\\Desktop\\x.xlsx","C:\\Users\\Developer\\Desktop\\x.csv",1);
            Console.WriteLine("ok");
        }
    }
}
