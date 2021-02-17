using System;

namespace Calculadora_facturas
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            ClsFile file = new ClsFile();
            file.cargarArchivo("D:\\Archivos\\28012021-Claro\\Factura Nov.xlsx");
            Console.ReadKey();

        }

    }
}
