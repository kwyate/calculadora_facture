using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Calculadora_factura_escritorio.Interfaces;


namespace Calculadora_factura_escritorio.Conexiones
{
    public class Conexion : IConexion
    {
        public OleDbConnection conexion(string conexionDriver, string nombreArchivo)
        {
            try
            {

                string conexion = string.Format(conexionDriver, nombreArchivo);

                OleDbConnection conector = new OleDbConnection(conexion);

                return conector;
            }
            catch (Exception e)
            {
                /*ProcessStartInfo info = new ProcessStartInfo();
                info.UseShellExecute = true;
                info.FileName = "AccessDatabaseEngine.exe";
                string aplicaticonPath = Application.StartupPath.Replace("bin\\Debug", "");
                string pagePath = Path.Combine(aplicaticonPath, "Resources");
                info.WorkingDirectory = pagePath;
                try
                {
                    Process.Start(info);

                    conexion(conexionDriver, nombreArchivo);
                }
                catch
                {
                    Debug.WriteLine("No se pudo ejecutar la aplicacion access");
                }*/
                
            }

            return null;
        }
    }
}
