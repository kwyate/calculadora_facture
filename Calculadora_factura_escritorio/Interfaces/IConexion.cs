using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculadora_factura_escritorio.Interfaces
{
    public interface IConexion
    {
        OleDbConnection conexion(string conexionDriver, string nombreArchivo);
    }
}
