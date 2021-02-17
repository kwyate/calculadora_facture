using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace Calculadora_factura_escritorio.Entidades
{
    class ObjConexion
    {
        public Conexiones.Conexion Conexion { get; set; }
        public Consultas.Query Query { get; set; }
        public OleDbConnection Conector { get; set; }
        public DataTableCollection dt { get; set; }
    }
}
