using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace MsExcelAutomatizacion
{
    class Program
    {
        static void Main(string[] args)
        {
            //Lectura de excel no standar

            FileExcelApp excel = new FileExcelApp(@"C:\repos\MsExcelUtility\PlantillaInformatica.xls");

            excel.Open();

            if(excel.getLevel() != "OK")
            {
                Console.WriteLine("Error {0}", excel.getMessage());
                Console.ReadKey();
                return;
            }
            string vWorksheetname = "Hoja1";

            if (!excel.existsWorkSheetName(vWorksheetname))
            {
                Console.WriteLine("La hoja no existe");
            }

            excel.setWorksheet(vWorksheetname);

            int rowStart = 6;
            //int rowFinish = excel.searchRowIndexInColRange("B","B","Rol");
            int rowFinish = excel.searchRowIndexInColRange("F", "F", "3");

            if(rowFinish > -1)
            {
                string begin = "G" + rowStart;
                string end = "J" + rowFinish;

                Console.WriteLine(begin);
                Console.WriteLine(end);
                string[][] values = excel.getCells(begin, end);
                foreach(string[] value in values)
                {
                    Console.WriteLine(value);
                }

            }
            else
            {
                Console.WriteLine("No se encontró un resultado de busqueda");
            }
            excel.Close();
            Console.ReadKey();
        }
    }
}
