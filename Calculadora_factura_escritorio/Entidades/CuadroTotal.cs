using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculadora_factura_escritorio.Entidades
{
    class CuadroTotal
    {
        public double GastoCelular { get; set; }
        public double iva { get; set; }
        public double imp { get; set; }
        public double reposicion { get; set; } = 0;
        public double otros_serv { get; set; } = 0;
        //Ajustes por revision de pagos
        public double ajus_rev_pag { get; set; } = 0;
        //Iva nuevos equipo
        public double iva_nue_equipos { get; set; } = 0;


    }
}
