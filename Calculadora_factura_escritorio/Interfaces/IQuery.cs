using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculadora_factura_escritorio.Interfaces
{
    interface IQuery
    {
        OleDbCommand queryConsulta(string query, OleDbConnection conector);
    }
}
