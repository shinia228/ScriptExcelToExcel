using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using ExcelDataReader.Exceptions;

namespace ScriptExcelToExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(@"C:\Users\User\Desktop\qwer\qwerty.xlsx");
            Excel.Worksheet worksheet = workbook.Sheets[1];

            Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet newWorksheet = newWorkbook.Sheets[1];

            int destRowNum = 1;
            int countInInterval = 0;

            for (int rowNum =4; rowNum <= 885; rowNum += 6) // изменение инкремента на 5
            {
                object value = worksheet.Cells[rowNum, 2].Value;

                if (value == null)
                {
                    break;
                }

                string cellValue = value.ToString();
                newWorksheet.Cells[destRowNum, 1].Value = cellValue;
                destRowNum++;

                countInInterval++;
            }


                newWorkbook.SaveAs(@"C:\Users\User\Desktop\qwer\test.xlsx");
                Console.WriteLine("Данные успешно записаны в новый файл Excel.");
            

            newWorkbook.Close();
            workbook.Close();

            excelApp.Quit();

            Console.WriteLine("Нажмите любую клавишу для выхода...");
            Console.ReadKey();
        }
    }
}

