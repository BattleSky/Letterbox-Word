using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LetterBox_Word.Letter;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using OfficeOpenXml;
using Template = LetterBox_Word.Letter.Template;


namespace LetterBox_Word
{


    class ExcelAsDbWork
    {

        // Метод обновления списка с перегрузкой
        public void CreateAndRenewDropdownItems(Corporation[] corporations, string chosenName = "")
        {
            // Генерация Списка корпораций
            var dropdown = Globals.Ribbons.Ribbon1.target_place;
            ClearDropdown(dropdown);

            foreach (var corp in corporations)
            {
                var item
                    = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                if (chosenName == "" || chosenName == "...")
                {
                    item.Label = corp.Name;
                    item.ScreenTip = corp.FullName;
                    item.SuperTip = corp.Adress;
                }

                dropdown.Items.Add(item);
            }
        }
        public void CreateAndRenewDropdownItems(PersonalNames[] names, string chosenName = "")
        {
            // Генерация Списка кому
            var dropdown = Globals.Ribbons.Ribbon1.target_person;
            ClearDropdown(dropdown);
            // Найти корпорацию по фамилии:

            foreach (var name in names)
            {
                var item
                    = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                if (chosenName == "" || chosenName == "...")
                {
                    item.Label = name.Name;
                    item.ScreenTip = name.Title;
                    item.SuperTip = name.CorporationName;
                }
                else
                {
                    if (name.CorporationName != chosenName) continue;
                    item.Label = name.Name;
                    item.ScreenTip = name.Title;
                    item.SuperTip = name.CorporationName;
                }

                dropdown.Items.Add(item);
            }
        }
        public void CreateAndRenewDropdownItems(Sender[] senders)
        {
            var dropdown = Globals.Ribbons.Ribbon1.from_name;
            ClearDropdown(dropdown);

            foreach (var sender in senders)
            {
                var item
                    = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = sender.Name;
                item.ScreenTip = sender.Title;
                dropdown.Items.Add(item);
            }
        }
        public void CreateAndRenewDropdownItems(Executor[] executors)
        {
            var dropdown = Globals.Ribbons.Ribbon1.creator_name;
            ClearDropdown(dropdown);

            foreach (var executor in executors)
            {
                var item
                    = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = executor.Name;
                dropdown.Items.Add(item);
            }
        }
        public void CreateAndRenewDropdownItems(Template[] templates)
        {
            var dropdown = Globals.Ribbons.Ribbon1.templates_dropDown;
            ClearDropdown(dropdown);
            foreach (var template in templates)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = template.Name;
                item.ScreenTip = template.InfoToInsert;
                dropdown.Items.Add(item);
            }
        }

        private void ClearDropdown(RibbonDropDown dropdown)
        {
            // Очистка
            dropdown.Items.Clear();
            var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item.Label = "...";
            dropdown.Items.Add(item);
        }

        public string[] GetDataOfOneUnit(Corporation[] corporations, string chosenName)
        {
            var result = new string[4];
            const string next = "\n";

            foreach (var corp in corporations)
            {
                if (corp.Name != chosenName) continue;
                var tel = "Тел: " + corp.Tel + next + "Факс: " + corp.Fax;
                if (corp.Tel == "")
                    tel = "Факс: " + corp.Fax;
                if (corp.Fax == "")
                    tel = "Тел: " + corp.Tel;
                if (corp.Tel == corp.Fax)
                    tel = "Тел/Факс: " + corp.Tel;

                var email = "";
                if (corp.Email != "")
                    email = "E-mail: " + corp.Email;

                result[0] = next + corp.FullName;
                result[1] = next + corp.Adress;
                result[2] = next + tel;
                result[3] = next + email;
            }
            return result;
        }
        public string[] GetDataOfOneUnit(PersonalNames[] names, string chosenName)
        {
            const string nextParagraph = "\n";
            var result = new string[3];
            foreach (var name in names)
            {
                if (name.Name != chosenName) continue;
                result[0] = name.Title;
                result[1] = nextParagraph + name.NameTo + nextParagraph;
                result[2] = "Уважаемый " + name.FirstMiddleName + "!";
            }
            return result;
        }
        public string[] GetDataOfOneUnit(Sender[] senders, string chosenName)
        {
            var result = new string[2];
            var regards = "С уважением,\n\n";
            foreach (var sender in senders)
            {
                if (sender.Name != chosenName) continue;
                result[0] = regards + sender.Title;
                result[1] = sender.NameToInsert;
            }
            return result;
        }
        public string[] GetDataOfOneUnit(Executor[] executors, string chosenName)
        {
            var result = new string[2];
            var isp = " Исп. ";
            var tel = "Тел. ";
            foreach (var executor in executors)
            {
                if (executor.Name != chosenName) continue;
                result[0] = isp + executor.Name + "\n";
                result[1] = tel + executor.Tel;
            }
            return result;
        }
        public string PreparingTextToInsert(Template[] templates, string chosenName)
        {
            string result = "";
            foreach (var template in templates)
            {
                if (template.Name != chosenName) continue;
                result = template.InfoToInsert;
            }
            return result;
        }


        #region Комментариев регион

        // На базе решения Excel Interop - рабочее решение, но проблема с пустыми ячейками

        //public string[] ReadFromExcel()
        //{
        //    var exlApp = new Microsoft.Office.Interop.Excel.Application();
        //    var exlWb = exlApp.Workbooks.Open(ExlFilePath, ReadOnly: false);
        //    exlApp.Visible = false; 
        //    var exlWs = exlWb.Worksheets[1] as Worksheet;
        //    var exlRange = exlWs.Range["B2"].Resize[100,1];
        //    var array = exlRange.Value;
        //    var resultList = new List<string>();
        //    for (int i = 1; i <= 14; i++)
        //    {
        //        var text = array[i, 1].Text.ToString();
        //        if (text != null)
        //            resultList.Add(text);
        //    }

        //    exlWb.Close(SaveChanges: false);
        //    exlApp.Quit();
        //    return resultList.ToArray();
        //}


        //private PersonalNames GetIndexesOfHeadingsNames(ExcelWorksheet sheet)
        //{
        //    var indexesOfHeadings = new PersonalNames();
        //    for (int column = 0; column < 20; column++)
        //    {
        //        switch (sheet.Cells[0, column].Text)
        //        {
        //            case "ФИО(Список)":
        //                PersonalNames.NameIndex = column;
        //                break;
        //            case "ФИО(Кому)":
        //                PersonalNames.NameToIndex = column;
        //                break;
        //            case "Должность(Кому)":
        //                PersonalNames.TitleIndex = column;
        //                break;
        //            case "Имя Отчество":
        //                PersonalNames.FirstMiddleNameIndex = column;
        //                break;
        //            case "Организация(Список)":
        //                PersonalNames.CorporationNameIndex = column;
        //                break;
        //        }

        //    }
        //    return indexesOfHeadings;
        //}


        //public string[] CollectDataFromColumn(int column)
        //{
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //    using (var workbookPackage = new ExcelPackage(new FileInfo(ExlFilePath)))
        //    {
        //        var resultList = new List<string>();
        //        var workbook = workbookPackage.Workbook;
        //        var sheet = workbook.Worksheets[0]; // 0-базовая (можно написать имя)
        //        for (var row = 1; row <= 100; row++)
        //        {
        //            //for (var column = 1; column <= 10; column++)
        //            //{
        //                var resultRow = sheet.Cells[row, column].Text;
        //                if (resultRow != "")
        //                    resultList.Add(resultRow);
        //            //}
        //        }

        //        return resultList.ToArray();
        //    }
        //}

        #endregion


    }
}
