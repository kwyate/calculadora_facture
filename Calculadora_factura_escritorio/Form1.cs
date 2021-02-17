using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Calculadora_factura_escritorio.Conexiones;
using Calculadora_factura_escritorio.Consultas;
using Calculadora_factura_escritorio.Entidades;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Calculadora_factura_escritorio.Acciones;
//using Application = Microsoft.Office.Interop.Excel.Application;

namespace Calculadora_factura_escritorio
{
    public partial class Form1 : Form
    {
        private Data data;
        List<Factura> listxls;
        List<DetallesCargos> listDetallesCargos;
        public Form1()
        {
            InitializeComponent();
            Iniciar.mostrarValores(cuadroGrid);
            data = new Data();

            //Importar_btn.Enabled = true;
        }

        public void activarForm()
        {
            txtGastosC.Enabled = true;
            txtImp.Enabled = true;
            txtIva.Enabled = true;
            txtOtro.Enabled = true;
            btnCalcular.Enabled = true;
        }
        //Limpia el campo cuando se cliquea por primera vez
        private void txtGastosC_Click_1(object sender, EventArgs e)
        {
            if (txtGastosC.Text.Equals("DIGITE EL VALOR"))
            {
                txtGastosC.Text = "";
                txtGastosC.ForeColor = Color.Black;
            }
        }
        //Limpia el campo cuando se cliquea por primera vez
        private void txtIva_Click(object sender, EventArgs e)
        {
            if (txtIva.Text.Equals("DIGITE EL VALOR"))
            {
                txtIva.Text = "";
                txtIva.ForeColor = Color.Black;
            }
        }
        //Limpia el campo cuando se cliquea por primera vez
        private void txtImp_Click(object sender, EventArgs e)
        {
            if (txtImp.Text.Equals("DIGITE EL VALOR"))
            {
                txtImp.Text = "";
                txtImp.ForeColor = Color.Black;
            }
        }
        private void txtOtro_Click(object sender, EventArgs e)
        {
            if (txtOtro.Text.Equals("DIGITE EL VALOR"))
            {
                txtOtro.Text = "";
                txtOtro.ForeColor = Color.Black;
            }
        }
        // Este evento se utiliza para validar que el txt tenga un valor de tipo double
        private void txtGastosC_KeyUp(object sender, KeyEventArgs e)
        {
            lblalertGS.Visible = !double.TryParse(txtGastosC.Text, out double n) ? true : false;
            GridViewObj.objCell(cuadroGrid, 0,2).Value = n.ToString();

        }
        // Este evento se utiliza para validar que el txt tenga un valor de tipo double
        private void txtIva_KeyUp(object sender, KeyEventArgs e)
        {
            lblAlertIva.Visible = !double.TryParse(txtIva.Text, out double n) ? true : false;
            GridViewObj.objCell(cuadroGrid, 1, 2).Value = n.ToString();

        }
        // Este evento se utiliza para validar que el txt tenga un valor de tipo double
        private void txtImp_KeyUp(object sender, KeyEventArgs e)
        {
            lblAlertImp.Visible = !double.TryParse(txtImp.Text, out double n) ? true : false;
            GridViewObj.objCell(cuadroGrid, 2, 2).Value = n.ToString();
        }
        // Este evento se utiliza para validar que el txt tenga un valor de tipo double
        private void txtOtro_KeyUp(object sender, KeyEventArgs e)
        {
            lblAlertaOtro.Visible = !double.TryParse(txtOtro.Text, out double n) ? true : false;
            GridViewObj.objCell(cuadroGrid, 7, 2).Value = n.ToString();

        }

        
        //--------------------------------------------------
        //Calcula las diferencias entre cuadro y factura
        //--------------------------------------------------
        private void btnCalcular_Click(object sender, EventArgs e)
        {
            GridViewObj.objCell(cuadroGrid, 8, 2).Value = Math.Round(double.Parse(txtGastosC.Text) + double.Parse(txtIva.Text) + double.Parse(txtImp.Text) + double.Parse(txtOtro.Text), 2);
            Diferencias.calcularDiferencias(cuadroGrid, new string[] {txtGastosC.Text, txtIva.Text, txtImp.Text});
            txtResumen.Text = data.resumen(GridViewObj.objCell(cuadroGrid, 8, 3).Value.ToString());

        }

        private void Importar_btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                Filter = "Excel | *.xls;*.xlsx;",
                Title = "Seleccionar Archivo"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Data data = new Data();
                txtFileName.Text = dialog.FileName;
                try
                {
                    this .listxls = Datos.obtenerDatos(dialog.FileName);
                    this.listDetallesCargos = Datos.DetallesCargos(listxls);

                    facturasGrid.DataSource = this.listDetallesCargos;

                    CuadroTotal DetallesFactura = Datos.DetallesFactura(listxls);

                    if (Validacion.cargarValoresCuadro(DetallesFactura, cuadroGrid)) activarForm();
                }
                catch
                {
                    string message = "No se encontro una hoja de calculo con el nombre de factura.\n Asegurese que la hoja a calcular tenga el nombre de factura";
                    string caption = "Error en la hoja de calculo";
                    MessageBoxButtons buttons = MessageBoxButtons.OK;
                    DialogResult result;

                    result = MessageBox.Show(message, caption, buttons);
                    if(result == System.Windows.Forms.DialogResult.OK)
                    {
                        this.Close();
                    }
                }

            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            var result = NewExcel.exportDocument(cuadroGrid, txtResumen.Text);
            if(result == 2)
            {
                string message = "El documento se descargo con exito lo encontrara en las descargas";
                string caption = "Exito de descarga";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult resultado;

                resultado = MessageBox.Show(message, caption, buttons);
                if (resultado == System.Windows.Forms.DialogResult.OK)
                {
                    this.Close();
                }
            }
        }

        private void facturasGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string numero_tel = GridViewObj.objCell(facturasGrid, e.RowIndex, 0).Value.ToString();


            var totalCalculos = Datos.listDetallesNumero(listxls, numero_tel);

            var numDetalles = Datos.numDetalles(listDetallesCargos, numero_tel);

            txt_numCel.Text = numDetalles.numero_tel;
            txtValorFac.Text = numDetalles.valor.ToString();
            txtIvaFac.Text = totalCalculos.Sum(x => x.iva).ToString();
            txtImpFac.Text = totalCalculos.Sum(x => x.imp).ToString();
            txtTotalFac.Text = numDetalles.total.ToString();

            facturaDetallesGrid.DataSource = Datos.valoresPositivos(totalCalculos);
            facturaDetallesDescuentos.DataSource = Datos.valoresDescuentos(totalCalculos);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnBuscarNum_Click(object sender, EventArgs e)
        {
            string numero_tel = txt_numCel.Text;

            var totalCalculos = Datos.listDetallesNumero(listxls, numero_tel);

            var numDetalles = Datos.numDetalles(listDetallesCargos, numero_tel);

            txt_numCel.Text = numDetalles.numero_tel;
            txtValorFac.Text = numDetalles.valor.ToString();
            txtIvaFac.Text = totalCalculos.Sum(x => x.iva).ToString();
            txtImpFac.Text = totalCalculos.Sum(x => x.imp).ToString();
            txtTotalFac.Text = numDetalles.total.ToString();


            facturaDetallesGrid.DataSource = Datos.valoresPositivos(totalCalculos);
            facturaDetallesDescuentos.DataSource = Datos.valoresDescuentos(totalCalculos);
        }

        
    }
}
