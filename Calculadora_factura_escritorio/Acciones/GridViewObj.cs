using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculadora_factura_escritorio.Acciones
{
    class GridViewObj
    {
        ///<summary>
        /// Trae el valor de un DataGridView de una celda en tipo string 
        ///</summary>
        public static double valCell(DataGridView table,int row, int cell) => double.Parse(objCell(table, row, cell).Value.ToString());
        ///<summary>
        ///obtiene la celda de un DataGridView
        ///</summary>
        public static DataGridViewCell objCell(DataGridView table, int row, int cell) => table.Rows[row].Cells[cell];

    }
}
