using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace LetterBox_Word.Letter
{
    class Executor
    {
        public string Name { get; private set; }
        public string Tel { get; private set; }
        public static int NameIndex { get; private set; }
        public static int TelIndex { get; private set; }

        private static string ExlFilePath => @"\\SERVERRAID\z\Помойки\Помойка_Степана\Programms\orion_forms\personalData.xlsx";

        public static void SetExecutorIndexes(ExcelWorksheet sheet)
        {
            for (int column = 1; column <= 20; column++)
            {
                switch (sheet.Cells[1, column].Text)
                {
                    case "ФИО":
                        NameIndex = column;
                        break;
                    case "Телефон":
                        TelIndex = column;
                        break;
                }
            }
        }
        public static Executor[] CollectExecutors
        {
            get
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var workbookPackage = new ExcelPackage(new FileInfo(ExlFilePath)))
                {
                    var executorList = new List<Executor>();
                    var workbook = workbookPackage.Workbook;
                    ExcelWorksheet sheet;
                    try
                    {
                        sheet = workbook.Worksheets["Исполнители"]; // 0-базовая (можно написать имя)
                        if (sheet == null)
                            throw new DirectoryNotFoundException("file:\n" + ExlFilePath + "not found");
                    }
                    catch (Exception e)
                    {
                        return new Executor[0];
                    }
                    Executor.SetExecutorIndexes(sheet);

                    for (var row = 2; row <= 100; row++) // Со второй сторки, первая - заголовки
                    {
                        var newExecutor = new Executor();
                        newExecutor.Name = sheet.Cells[row, Executor.NameIndex].Text;
                        if (newExecutor.Name == "") break;
                        newExecutor.Tel = sheet.Cells[row, Executor.TelIndex].Text;
                        executorList.Add(newExecutor);
                    }
                    return executorList.ToArray();
                }
            }
        }
    }
}
