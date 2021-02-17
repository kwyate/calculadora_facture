using Calculadora_factura_escritorio.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculadora_factura_escritorio.Consultas
{
    class Query : IQuery
    {
        public OleDbCommand queryConsulta(string query, OleDbConnection conector)
        {
            try
            {
                OleDbCommand consulta = new OleDbCommand(query, conector);
                return consulta;
            }
            catch (Exception e)
            {
                Console.WriteLine("Hay un error: " + e);
            }
            return null;
        }
    }
}
