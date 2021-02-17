using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculadora_factura_escritorio.Entidades
{
    class FacturaDetalles
    {
        public string numero { get; set; }//Numero Celular
        public string nombre { get; set; }// Nombre empresa
        public string descripcion { get; set; }
        public double valor { get; set; }
        public double iva { get; set; }
        public double imp { get; set; }
        public double total { get; set; }
    }
}
