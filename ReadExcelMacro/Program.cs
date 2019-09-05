using ClosedXML.Excel;
using System;


namespace ReadExcelMacro
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Leitura de arquivo EXCEL(XLSM).");

            string fileName = @"C:\Users\leonardo.queres\source\repos\ReadExcelMacro\ReadExcelMacro\Excel\excel.xlsm";

            LerExcel(fileName);
        }

        private static void LerExcel(string fileName)
        {
            using (var wb = new XLWorkbook(fileName))
            {
                var ws = wb.Worksheet(1).RangeUsed();

                // Look for the first row used
                var firstRowUsed = ws.FirstRowUsed();

                // Narrow down the row so that it only includes the used part
                var row = firstRowUsed.RowUsed();

                // Move to the next row (it now has the titles)
                //categoryRow = categoryRow.RowBelow();

                var cont = 1;
                var totalLinha = 1;
                while (!row.IsEmpty())
                {
                    Console.WriteLine(totalLinha++);
                    while (cont <= 23)
                    {
                        var categoryName = row.Cell(cont).GetString();

                        Console.WriteLine(cont + " - " + categoryName);
                        cont++;
                    }

                    row = row.RowBelow();
                    cont = 1;


                    Console.WriteLine("--------------------------------------------------------------------");
                    Console.WriteLine("--------------------------------------------------------------------");
                    Console.WriteLine("--------------------------------------------------------------------");
                    Console.WriteLine("--------------------------------------------------------------------");
                }
                Console.ReadLine();
            }
        }
    }
}
