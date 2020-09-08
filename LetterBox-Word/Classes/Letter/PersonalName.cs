using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace LetterBox_Word.Letter
{
    public class PersonalNames
    {
        public static int NameIndex { get; private set; }
        public static int NameToIndex { get; private set; }
        public static int FirstMiddleNameIndex { get; private set; }
        public static int TitleIndex { get; private set; }
        public static int CorporationNameIndex { get; private set; }

        public string Name { get; private set; }
        public string NameTo { get; private set; }
        public string FirstMiddleName { get; private set; }
        public string Title { get; private set; }
        public string CorporationName { get; private set; }

        private static string  ExlFilePath => @"\\SERVERRAID\z\Помойки\Помойка_Степана\Programms\orion_forms\personalData.xlsx";
        public static void SetPersonalIndexes(ExcelWorksheet sheet)
        {
            for (int column = 1; column <= 20; column++)
            {
                switch (sheet.Cells[1, column].Text)
                {
                    case "ФИО(Список)":
                        NameIndex = column;
                        break;
                    case "ФИО(Кому)":
                        NameToIndex = column;
                        break;
                    case "Должность(Кому)":
                        TitleIndex = column;
                        break;
                    case "Имя Отчество":
                        FirstMiddleNameIndex = column;
                        break;
                    case "Организация(Список)":
                        CorporationNameIndex = column;
                        break;
                }

            }
        }

        public static PersonalNames[] CollectPersonNames
        {
            get
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var workbookPackage = new ExcelPackage(new FileInfo(ExlFilePath)))
                {
                    var personalDataList = new List<PersonalNames>();
                    var workbook = workbookPackage.Workbook;
                    ExcelWorksheet sheet;
                    try
                    {
                        sheet = workbook.Worksheets["Имена"]; // 0-базовая (можно написать имя)
                        if (sheet == null)
                            throw new DirectoryNotFoundException("file:\n" + ExlFilePath + "not found");
                    }
                    catch (Exception e)
                    {
                        return new PersonalNames[0];
                    }

                    PersonalNames.SetPersonalIndexes(sheet);

                    for (var row = 2; row <= 100; row++) // Со второй сторки, первая - заголовки
                    {
                        var newPerson = new PersonalNames();
                        newPerson.Name = sheet.Cells[row, PersonalNames.NameIndex].Text;
                        if (newPerson.Name == "") break;
                        newPerson.NameTo = sheet.Cells[row, PersonalNames.NameToIndex].Text;
                        newPerson.FirstMiddleName = sheet.Cells[row, PersonalNames.FirstMiddleNameIndex].Text;
                        newPerson.CorporationName = sheet.Cells[row, PersonalNames.CorporationNameIndex].Text;
                        newPerson.Title = sheet.Cells[row, PersonalNames.TitleIndex].Text;
                        personalDataList.Add(newPerson);
                    }

                    return personalDataList.ToArray();
                }
            }
        }
    }
}
