using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculadora_factura_escritorio.Entidades
{
    class DetallesCargos
    {
        public string numero_tel { get; set; }
        public double valor { get; set; } 
        public double impuesto { get; set; }
        public double total { get; set; }
    }
}
