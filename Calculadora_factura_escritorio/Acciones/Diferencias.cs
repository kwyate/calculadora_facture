using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculadora_factura_escritorio.Acciones
{
    class Diferencias
    {

        public static void calcularDiferencias(DataGridView table, string [] text)
        {
            //calcular diferencia Gasto Celular = factura_GS -(cuadroGs + CuadroReposicion + cuadroOtrosServicios + CuadroAjustexRev)
            GridViewObj.objCell(table, 0, 3).Value = Math.Round(double.Parse(text[0]) - (GridViewObj.valCell(table, 0, 1) + GridViewObj.valCell(table, 3, 1) + GridViewObj.valCell(table, 4, 1) + GridViewObj.valCell(table,5, 1)), 2);
            //Calcular diferencia IVA = factura_IVA -(CuadroIva + cuadroIvaNuevosEquipos)
            GridViewObj.objCell(table, 1, 3).Value = Math.Round(double.Parse(text[1]) - (GridViewObj.valCell(table, 1, 1) + GridViewObj.valCell(table, 6, 1)), 2);
            //Calcular diferencia IMP = factura_Imp - CuadroIMP
            GridViewObj.objCell(table, 2, 3).Value = Math.Round(double.Parse(text[2]) - GridViewObj.valCell(table, 2, 1), 2);
            //Calcular total diferencia = la sumatoria de los anteriores calculos
            GridViewObj.objCell(table, 8, 3).Value = Math.Round(GridViewObj.valCell(table, 0, 3) + GridViewObj.valCell(table, 1, 3) + GridViewObj.valCell(table,2, 3), 2);
        }
    }
}
