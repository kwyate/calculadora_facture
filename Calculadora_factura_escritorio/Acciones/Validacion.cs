using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Calculadora_factura_escritorio.Entidades;

namespace Calculadora_factura_escritorio.Acciones
{
    class Validacion
    {
        ///<summary>
        /// Devuelve un valor booleano despues de que cargue los datos del archivo xls en el cuadro de diferencias
        ///</summary>
        public static bool cargarValoresCuadro(CuadroTotal dt, DataGridView table)
        {
            try
            {
                GridViewObj.objCell(table, 0, 1).Value = dt.GastoCelular.ToString();
                GridViewObj.objCell(table, 1, 1).Value = dt.iva.ToString();
                GridViewObj.objCell(table, 2, 1).Value = dt.imp.ToString();
                GridViewObj.objCell(table, 8, 1).Value = (dt.GastoCelular + dt.imp + dt.iva).ToString();
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
