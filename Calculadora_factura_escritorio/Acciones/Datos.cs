using Calculadora_factura_escritorio.Entidades;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Calculadora_factura_escritorio.Consultas;

namespace Calculadora_factura_escritorio.Acciones
{
    class Datos
    {
        ///<summary>
         /// Obtiene los datos del excel y los convierte en un objeto JSON
         ///</summary>
        public static List<Factura> obtenerDatos(string nombreArchivo)
        {
            string query = "SELECT MAESTRA AS maestra," +
                    " CUSTCODE AS custcode," +
                    " MIN AS numero, " +
                    " NOMBRE AS nombre," +
                    " DESCRIPCION AS descripcion," +
                    " VALOR AS valor," +
                    " IVA AS iva," +
                    " CICLO AS ciclo " +
                    " FROM [factura$] WHERE maestra <> ''";
            DataTableCollection detallesCargosnums = Data.ImportDataView(nombreArchivo, query);
            string JSONString = string.Empty;
            JSONString = JsonConvert.SerializeObject(detallesCargosnums[0]);
            return JsonConvert.DeserializeObject<List<Factura>>(JSONString);
        }
        public static List<DetallesCargos> DetallesCargos(List<Factura> facturas)
        {
            return facturas.GroupBy(x => x.numero).OrderByDescending(x => x.Key).Select(x => new DetallesCargos
            {
                numero_tel = x.First().numero.Split('.')[0],
                valor = Math.Round(x.Sum(v => v.valor), 2),
                impuesto = Math.Round(x.Sum(i => i.iva), 2),
                total = Math.Round(x.Sum(t => t.valor) + x.Sum(t => t.iva), 2)

            }).ToList<DetallesCargos>();
        }
        public static CuadroTotal DetallesFactura(List<Factura> facturas)
        {
            return facturas.Select(x => new CuadroTotal
            {
                GastoCelular = Math.Round((from v in facturas select v.valor).Sum(), 2),
                iva = Math.Round((from v in facturas select Math.Round(v.valor * 0.19, 2)).Sum(), 2), //
                imp = Math.Round((from v in facturas select Math.Round(v.valor * 0.19, 2) == v.iva ? 0 : Math.Round(v.valor * 0.04, 2)).Sum(), 2),//
                reposicion = 0,
                otros_serv = 0,
                ajus_rev_pag = 0,
                iva_nue_equipos = 0
            }).FirstOrDefault();
        }
        public static List<FacturaDetalles> listDetallesNumero(List<Factura> facturas, string numero_tel)
        {
            //Se calculan primero los valores iva e imp
            var listDetallesFactura = facturas.Where(x => x.numero.Split('.')[0].Equals(numero_tel)).Select(x => new FacturaDetalles
            {
                numero = x.numero,
                nombre = x.nombre,
                descripcion = x.descripcion,
                iva = Math.Round(x.valor * 0.19, 2),
                imp = Math.Round(x.valor * 0.19, 2) == x.iva ? 0 : Math.Round(x.valor * 0.04, 2),
                valor = x.valor,
                total = 0

            }).ToList<FacturaDetalles>();
            //Despues de calculados los anteriores valores se calcula el total
            return listDetallesFactura.Select(x => new FacturaDetalles
            {
                numero = x.numero.Split('.')[0],
                nombre = x.nombre,
                descripcion = x.descripcion,
                iva = x.iva,
                imp = x.imp,
                valor = x.valor,
                total = Math.Round(x.valor + x.iva + x.imp, 2)
            }).ToList<FacturaDetalles>();
        }
        public static DetallesCargos numDetalles(List<DetallesCargos> dc, string numero_tel) => dc.Where(x => x.numero_tel.Equals(numero_tel)).FirstOrDefault();
        public static List<FacturaDetalles> valoresPositivos(List<FacturaDetalles> facturaDetalles) => facturaDetalles.Where(x => x.valor > -1).ToList<FacturaDetalles>();
        public static List<FacturaDetalles> valoresDescuentos(List<FacturaDetalles> facturaDetalles) => facturaDetalles.Where(x => x.valor < 0).ToList<FacturaDetalles>();

    }
}
