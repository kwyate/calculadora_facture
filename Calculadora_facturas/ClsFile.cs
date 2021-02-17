using ExcelDataReader;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
namespace Calculadora_facturas
{
    class ClsFile
    {

        public DataView ImportarDatos(string nombreArchivo)
        {
            string conexion = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0'");

            OleDbConnection conector = new OleDbConnection(conexion);

            conector.Open();

            OleDbCommand consulta = new OleDbCommand("select from [Hoja1$]", conector) ;


            OleDbDataAdapter adaptador = new OleDbDataAdapter()
            {
                SelectCommand = consulta
            };

            DataSet ds = new DataSet();

            adaptador.Fill(ds);
            conector.Close();

            return ds.Tables[0].DefaultView;

        }

        





        public void cargarArchivo(String filePath)
        {
            
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                {
                    
                    FallbackEncoding = Encoding.GetEncoding(1252),

                    // Gets or sets the password used to open password protected workbooks.
                    //--Password = "password",

                    // Gets or sets an array of CSV separator candidates. The reader 
                    // autodetects which best fits the input data. Default: , ; TAB | # 
                    // (CSV only)
                    //-AutodetectSeparators = new char[] { ',', ';', '\t', '|', '#' },

                    // Gets or sets a value indicating whether to leave the stream open after
                    // the IExcelDataReader object is disposed. Default: false
                    //-LeaveOpen = false,

                    // Gets or sets a value indicating the number of rows to analyze for
                    // encoding, separator and field count in a CSV. When set, this option
                    // causes the IExcelDataReader.RowCount property to throw an exception.
                    // Default: 0 - analyzes the entire file (CSV only, has no effect on other
                    // formats)
                    //-AnalyzeInitialCsvRows = 0,
                });
                
                    do
                    {
                        while (reader.Read())
                        {
                            reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

            }
        }


    }
}
