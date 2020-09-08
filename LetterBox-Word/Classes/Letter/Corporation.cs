using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace LetterBox_Word.Letter
{
    public class Corporation
    {
        public static int NameIndex { get; private set; }
        public static int FullNameIndex { get; private set; }
        public static int AdressIndex { get; private set; }
        public static int TelIndex { get; private set; }
        public static int FaxIndex { get; private set; }
        public static int EmailIndex { get; private set; }

        public string Name { get; private set; }
        public string FullName { get; private set; }
        public string Adress { get; private set; }
        public string Tel { get; private set; }
        public string Fax { get; private set; }
        public string Email { get; private set; }

        private static string ExlFilePath => @"\\SERVERRAID\z\Помойки\Помойка_Степана\Programms\orion_forms\personalData.xlsx";

        public static void SetCorporationIndexes(ExcelWorksheet sheet)
        {
            for (int column = 1; column <= 20; column++)
            {
                switch (sheet.Cells[1, column].Text)
                {
                    case "Наименование(Список)":
                        NameIndex = column;
                        break;
                    case "Наименование(Полное)":
                        FullNameIndex = column;
                        break;
                    case "Адрес":
                        AdressIndex = column;
                        break;
                    case "Тел":
                        TelIndex = column;
                        break;
                    case "Факс":
                        FaxIndex = column;
                        break;
                    case "E-Mail":
                        EmailIndex = column;
                        break;
                }

            }
        }

        public static Corporation[] CollectCorporations
        {
            get
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var workbookPackage = new ExcelPackage(new FileInfo(ExlFilePath)))
                {
                    var corporationList = new List<Corporation>();
                    var workbook = workbookPackage.Workbook;
                    ExcelWorksheet sheet;
                    try
                    {
                        sheet = workbook.Worksheets["Организации"]; // 0-базовая (можно написать имя)
                        if (sheet == null)
                            throw new DirectoryNotFoundException("file:\n" + ExlFilePath + "not found");
                    }
                    catch (Exception e)
                    {
                        return new Corporation[0];
                    }
                    Corporation.SetCorporationIndexes(sheet);

                    for (var row = 2; row <= 100; row++) // Со второй сторки, первая - заголовки
                    {
                        var newCorp = new Corporation();
                        newCorp.Name = sheet.Cells[row, Corporation.NameIndex].Text;
                        if (newCorp.Name == "") break;
                        newCorp.FullName = sheet.Cells[row, Corporation.FullNameIndex].Text;
                        newCorp.Adress = sheet.Cells[row, Corporation.AdressIndex].Text;
                        newCorp.Email = sheet.Cells[row, Corporation.EmailIndex].Text;
                        newCorp.Tel = sheet.Cells[row, Corporation.TelIndex].Text;
                        newCorp.Fax = sheet.Cells[row, Corporation.FaxIndex].Text;
                        corporationList.Add(newCorp);
                    }

                    return corporationList.ToArray();
                }
            }
        }
    }
}
