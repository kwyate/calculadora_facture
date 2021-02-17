using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Calculadora_factura_escritorio.Acciones
{
    class NewExcel
    {
        ///<summary>
        /// Agrega estilos a las celdas del nuevo excel
        ///</summary>
        private static void styleCells(Workbook libro)
        {
            Func<string, Range> range = (r) => libro.Sheets[1].Range(r);
            range("A1").Interior.ColorIndex = 14;
            range("B1").Interior.ColorIndex = 14;
            range("C1").Interior.ColorIndex = 47;
            range("D1").Interior.ColorIndex = 53;
            range("A1:D9").Borders.LineStyle = XlLineStyle.xlDouble;
            range("A1:D9").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range("A1:D15").Font.Size = 12;
            range("A1:D1").Font.Bold = true;
            range("A1:A9").Font.Bold = true;
            range("A1").ColumnWidth = 50;
            range("B1:D9").ColumnWidth = 25;
            range("A11:D15").MergeCells = true;
            range("A11:D15").Font.Bold = true;
            range("A11:D15").WrapText = true;
            range("A11:D15").Borders.LineStyle = XlLineStyle.xlContinuous;
        }
        ///<summary>
        /// Exporta el nuevo excel
        ///</summary>
        private static void exportalExcel(DataGridView tb, string resumen)
        {
            //Se crea la instancia de aplicacion excel
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //Se agrega una hoja al excel
            Workbook libro = excel.Workbooks.Add(true);

            for (int i = 0; i < tb.Columns.Count; i++)
                libro.Sheets[1].Cells[1, i + 1] = tb.Columns[i].HeaderText;


            for (int f = 0; f < tb.Rows.Count; f++)
                for (int c = 0; c < tb.Columns.Count; c++)
                    libro.Sheets[1].Cells[f + 2, c + 1] = tb.Rows[f].Cells[c].Value;

            libro.Sheets[1].Cells[11, 1] = resumen;
            styleCells(libro);

            excel.Visible = true;

        }
        ///<summary>
        /// Agrega estilos a las celdas del nuevo excel
        ///</summary>
        private static void styleCellsEPPlus(ExcelWorksheet libro)
        {
            Func<string, ExcelRange> range = (r) => libro.Cells[r];
            range("A1:D1").Style.Fill.PatternType = ExcelFillStyle.Solid;
            range("A1").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSkyBlue);
            range("B1").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            range("C1").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSlateGray);
            range("D1").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkSeaGreen);
            range("A1:D10").Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range("A1:D10").Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range("A1:D10").Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range("A1:D10").Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range("A1:D10").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range("A1:D15").Style.Font.Size = 12;
            range("A1:D1").Style.Font.Bold = true;
            range("A1:A9").Style.Font.Bold = true;
            range("A1").AutoFitColumns(100);
            range("B1:D9").AutoFitColumns(100);
            range("A11:D15").Merge = true;
            range("A11:D15").Style.Font.Bold = true;
            range("A11:D15").Style.WrapText = true;
            range("A11:D15").Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range("A11:D15").Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range("A11:D15").Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range("A11:D15").Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }
        private static void exportarExcelEPPlus(DataGridView tb, string resumen)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                var libro = excel.Workbook.Worksheets["Worksheet1"];
                for (int i = 0; i < tb.Columns.Count; i++)
                    libro.Cells[1, i+1].Value =  tb.Columns[i].HeaderText; 


                for (int f = 0; f < tb.Rows.Count; f++)
                    for (int c = 0; c < tb.Columns.Count; c++)
                        libro.Cells[f + 2, c + 1].Value = tb.Rows[f].Cells[c].Value;

                libro.Cells[11, 1].Value = resumen;
                styleCellsEPPlus(libro);
                string username = Environment.UserName;
                FileInfo excelFile = new FileInfo(@"C:\Users\"+username+@"\Downloads\CuadroDiferencia.xlsx");
                excel.SaveAs(excelFile);
            }
        }
        public static int exportDocument(DataGridView tb, string resumen)
        {

            try
            {
                exportalExcel(tb, resumen);
                return 1;
            }
            catch
            {
                exportarExcelEPPlus(tb, resumen);
                return 2;
            }
        
        }
    }
}
