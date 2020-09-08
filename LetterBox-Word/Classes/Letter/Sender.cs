using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace LetterBox_Word.Letter
{
    class Sender
    {
        public string Name { get; private set; }
        public string NameToInsert { get; private set; }
        public string Title { get; private set; }

        public static int NameIndex { get; private set; }
        public static int NameToInsertIndex { get; private set; }
        public static int TitleIndex { get; private set; }
        private static string ExlFilePath => @"\\SERVERRAID\z\Помойки\Помойка_Степана\Programms\orion_forms\personalData.xlsx";

        public static void SetSenderIndexes(ExcelWorksheet sheet)
        {
            for (int column = 1; column <= 20; column++)
            {
                switch (sheet.Cells[1, column].Text)
                {
                    case "ФИО(Список)":
                        NameIndex = column;
                        break;
                    case "Должность":
                        TitleIndex = column;
                        break;
                    case "ФИО":
                        NameToInsertIndex = column;
                        break;
                }

            }
        }

        public static Sender[] CollectSenders
        {
            get
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var workbookPackage = new ExcelPackage(new FileInfo(ExlFilePath)))
                {
                    var senderList = new List<Sender>();
                    var workbook = workbookPackage.Workbook;
                    ExcelWorksheet sheet;
                    try
                    {
                        sheet = workbook.Worksheets["Адресанты"]; // 0-базовая (можно написать имя)
                        if (sheet == null)
                            throw new DirectoryNotFoundException("file:\n" + ExlFilePath + "not found");
                    }
                    catch (Exception e)
                    {
                        return new Sender[0];
                    }
                    Sender.SetSenderIndexes(sheet);

                    for (var row = 2; row <= 100; row++) // Со второй сторки, первая - заголовки
                    {
                        var newSender = new Sender();
                        newSender.Name = sheet.Cells[row, Sender.NameIndex].Text;
                        if (newSender.Name == "") break;
                        newSender.NameToInsert = sheet.Cells[row, Sender.NameToInsertIndex].Text;
                        newSender.Title = sheet.Cells[row, Sender.TitleIndex].Text;
                        senderList.Add(newSender);
                    }
                    return senderList.ToArray();
                }
            }
        }

    }
}
