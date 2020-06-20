using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CreateOandT
{
    class Program
    {
        static void Main(string[] args)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.ScreenUpdating = false;
            ExcelApp.Visible = false;
            ExcelApp.Interactive = false;
            ExcelApp.IgnoreRemoteRequests = true;


            string fName = @"C:\Users\35498\source\repos\DataSetExcel\Neuro\Vse_dannye.xlsx"; // Файл Excel, с которым производится работа
            string sheetName = @"DataResult";                 // Название листа откуда берётся информация
            string sheetTeachName = @"О";     // Название листа откуда куда поместятся преобразованные данные
            string sheetTestName = @"Т"; // Название листа где будет обозначение данных

            var wb = ExcelApp.Workbooks.Open(fName);

            try
            {
                var sheet = (Excel.Worksheet)wb.Worksheets[sheetName];
                Excel.Worksheet teachSheet;
                Excel.Worksheet testSheet;

                gg1:
                try
                {
                    teachSheet = (Excel.Worksheet)wb.Worksheets[sheetTeachName];
                    sheetTeachName += "1";
                    goto gg1;
                }
                catch { }

                gg2:
                try
                {
                    testSheet = (Excel.Worksheet)wb.Worksheets[sheetTestName];
                    sheetTestName += "1";
                    goto gg2;
                }
                catch { }

                ((Excel.Worksheet)wb.Worksheets.Add()).Name = sheetTeachName;
                teachSheet = (Excel.Worksheet)wb.Worksheets[sheetTeachName];
                ((Excel.Worksheet)wb.Worksheets.Add()).Name = sheetTestName;
                testSheet = (Excel.Worksheet)wb.Worksheets[sheetTestName];


                int columnIndex = 1;
                int countOfRow = 0;
                while (sheet.Cells[countOfRow + 1, 1].Value != null) countOfRow++;
                List<int> teachSet = new List<int>();
                for (int i = 2; i <= countOfRow + 1; i++)
                    teachSet.Add(i);
                List<int> testSet = new List<int>();

                var rnd = new Random();
                for (int i = 0; i < (int)(countOfRow * 0.2); i++)
                {
                    int randInt = rnd.Next(1, countOfRow);
                    if (testSet.Contains(teachSet[randInt]))
                    {
                        i--;
                        continue;
                    }
                    testSet.Add(teachSet[randInt]);
                }
                foreach (var item in testSet) teachSet.Remove(item);

                int counOfColumn = 1;
                while (sheet.Cells[1, counOfColumn].Value != null)
                {
                    teachSheet.Cells[1, counOfColumn] = "X" + counOfColumn.ToString();
                    testSheet.Cells[1, counOfColumn] = "X" + counOfColumn.ToString();
                    counOfColumn++;
                }
                teachSheet.Cells[1, counOfColumn - 1] = "D" + 1.ToString();
                testSheet.Cells[1, counOfColumn - 1] = "D" + 1.ToString();
                counOfColumn = counOfColumn - 1;

                int to = 2;
                foreach (int i in teachSet)
                {
                    for (int j = 1; j <= counOfColumn; j++)
                    {
                        teachSheet.Cells[to, j] = sheet.Cells[i, j];
                    }
                    to++;
                }
                int g = GetColumnIndex(sheet, "Время");
                teachSheet.Range[teachSheet.Cells[2, 1], teachSheet.Cells[to - 1, counOfColumn]].NumberFormat = "0";
                teachSheet.Range[teachSheet.Cells[2, g], teachSheet.Cells[to - 1, g]].NumberFormat = "0,00";

                to = 2;
                foreach (int i in testSet)
                {
                    for (int j = 1; j <= counOfColumn; j++)
                    {
                        testSheet.Cells[to, j] = sheet.Cells[i, j];
                    }
                    to++;
                }
                testSheet.Range[testSheet.Cells[2, 1], testSheet.Cells[to - 1, counOfColumn]].NumberFormat = "0";
                testSheet.Range[testSheet.Cells[2, g], testSheet.Cells[to - 1, g]].NumberFormat = "0,00";


                testSheet.Columns.EntireColumn.AutoFit();
                teachSheet.Columns.EntireColumn.AutoFit();

                // Console.WriteLine("Всего " + i);
            }
            catch (Exception e)
            {
                Console.WriteLine("!!!!");
                Console.WriteLine(e.Message);
            }


            ExcelApp.ScreenUpdating = true;
            ExcelApp.Interactive = true;
            ExcelApp.IgnoreRemoteRequests = false;
            ExcelApp.Visible = true;
            Console.ReadLine();
        }
        private static int GetColumnIndex(Excel.Worksheet sheet, string columnName)
        {
            int clmnIndx = -1;
            for (int i = 1; sheet.Cells[1, i].Value != null; i++)
            {
                if (sheet.Cells[1, i].Text == columnName)
                {
                    clmnIndx = i;
                }
            }
            return clmnIndx;
        }
    }
}
