using Calculadora_factura_escritorio.Conexiones;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculadora_factura_escritorio.Consultas
{
    public class Data
    {
        ///<summary>
        /// Importa datos
        ///</summary>
        public static DataTableCollection ImportDataView(string nombreArchivo, string query)
        {
            Conexion c = new Conexion();
            Query q = new Query();
            var conector = c.conexion("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties='Excel 12.0 Xml; HDR = YES'", nombreArchivo);

                conector.Open();
           
            var consulta = q.queryConsulta(query, conector);
            OleDbDataAdapter adaptador = new OleDbDataAdapter()
            {
                SelectCommand = consulta
            };
            DataSet ds = new DataSet();
            adaptador.Fill(ds);
            conector.Close();
            return ds.Tables;
        }
        ///<summary>
        /// Resumen de las dierencias de factura y cuadro
        ///</summary>
        public string resumen(string val1)
        {
            return "LA DIFERENCIA ENTRE EL XLS Y LA FACTURA TIENE UN VALOR DE : $"+val1+".";
        }
    }
}
