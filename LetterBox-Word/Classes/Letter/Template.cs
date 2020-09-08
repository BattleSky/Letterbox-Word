using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace LetterBox_Word.Letter
{
    class Template
    {
        public string Name { get; private set; }
        public string InfoToInsert { get; private set; }
        public static int NameIndex { get; private set; }
        public static int InfoToInsertIndex { get; private set; }

        private static string ExlFilePath => @"\\SERVERRAID\z\Помойки\Помойка_Степана\Programms\orion_forms\personalData.xlsx";

        public static void SetTemplateIndexes(ExcelWorksheet sheet)
        {
            for (int column = 1; column <= 20; column++)
            {
                switch (sheet.Cells[1, column].Text)
                {
                    case "Название(Кратко)":
                        NameIndex = column;
                        break;
                    case "Данные для вставки":
                        InfoToInsertIndex = column;
                        break;
                }
            }
        }
        public static Template[] CollectTemplates
        {
            get
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var workbookPackage = new ExcelPackage(new FileInfo(ExlFilePath)))
                {
                    var templateList = new List<Template>();
                    var workbook = workbookPackage.Workbook;
                    ExcelWorksheet sheet;
                    try
                    {
                        sheet = workbook.Worksheets["Шаблоны"]; // 0-базовая (или имя)
                        if (sheet == null)
                            throw new DirectoryNotFoundException();
                    }
                    catch
                    {
                        var text = "Ошибка при обращении к одной из страниц файла данных. Проверьте\n" +
                                   "*Есть ли доступ к серверу SERVERRAID?\n" +
                                   "*Доступен ли файл по адресу:\n" + ExlFilePath + "?\n" +
                                   "\n\n";
                                    MessageBox.Show(text);
                        return new Template[0];
                    }
                    Template.SetTemplateIndexes(sheet);

                    for (var row = 2; row <= 100; row++) // Со второй сторки, первая - заголовки
                    {
                        var newTemplate = new Template();
                        newTemplate.Name = sheet.Cells[row, Template.NameIndex].Text;
                        if (newTemplate.Name == "") break;
                        newTemplate.InfoToInsert = sheet.Cells[row, Template.InfoToInsertIndex].Text;
                        templateList.Add(newTemplate);
                    }
                    return templateList.ToArray();
                }
            }
        }

    }
    

}
