using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculadora_factura_escritorio.Acciones
{
    class Iniciar
    {
        public static void mostrarValores(DataGridView table)
        {
            string[] row = new string[] { "Gasto Celular", "0", "0", "0" };
            table.Rows.Add(row);
            row = new string[] { "IVA 19%", "0", "0", "0" };
            table.Rows.Add(row);
            row = new string[] { "Imp. Consumo 4%", "0", "0", "0" };
            table.Rows.Add(row);
            row = new string[] { "Reposición", "0", "", "" };
            table.Rows.Add(row);
            row = new string[] { "Otros Servicios y Créditos", "0", "", "" };
            table.Rows.Add(row);
            row = new string[] { "Ajuste x Rev de pagos", "0", "", "" };
            table.Rows.Add(row);
            row = new string[] { "IVA nuevos equipos", "0", "", "" };
            table.Rows.Add(row);
            row = new string[] { "Otros Factura", "", "0", "" };
            table.Rows.Add(row);
            row = new string[] { "Total", "0", "0", "0" };
            table.Rows.Add(row);
        }
    }
}
