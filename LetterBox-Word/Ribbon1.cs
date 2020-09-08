using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Sockets;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using LetterBox_Word.Letter;
using Microsoft.Office.Tools.Ribbon;
using OfficeOpenXml.Drawing.Style.Effect;
using Word = Microsoft.Office.Interop.Word;


namespace LetterBox_Word
{
    public partial class Ribbon1
    {
        private readonly ExcelAsDbWork excelFile = new ExcelAsDbWork();
        private readonly PersonalNames[] names = PersonalNames.CollectPersonNames;
        private readonly Corporation[] corporations = Corporation.CollectCorporations;
        private readonly Sender[] senders = Sender.CollectSenders;
        private readonly Executor[] executors = Executor.CollectExecutors;
        private readonly Template[] templates = Template.CollectTemplates;
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            excelFile.CreateAndRenewDropdownItems(templates);
        }

        #region Хранение статуса галочек и дропбокса о письме

        private LetterInfo firstAddressInfo = new LetterInfo()
        {
            AddressCheckBox = true, EmailCheckBox = true, 
            IsACopyCheckBox = false, TelFaxCheckBox = true
        };
        private LetterInfo secondAddressInfo = new LetterInfo()
        {
            AddressCheckBox = true, EmailCheckBox = true, 
            IsACopyCheckBox = false, TelFaxCheckBox = true
        };
        private LetterInfo thirdAddressInfo = new LetterInfo()
        {
            AddressCheckBox = true, EmailCheckBox = true,
            IsACopyCheckBox = false, TelFaxCheckBox = true
        };


        #endregion

        #region Кнопки управления письмом

        private void btnLetterTemp_Click(object sender, RibbonControlEventArgs e)
        {
            object missing = System.Type.Missing;
            try
            {
                Globals.ThisAddIn.Application.Documents.Add(
                    @"\\SERVERRAID\z\Помойки\Помойка_Степана\Programms\orion_forms\letter_form.docx", ref missing,
                    ref missing, ref missing);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Ошибка при обращении к файлу формы:\n" + exception.Message);
            }

            Letter_Box.Visible = true;
            turnLetterFormOnOff.Checked = true;
            btn_Address1.Checked = chk_telfax.Checked = chk_address.Checked = chk_address.Checked = true;
            btn_Address2.Checked = btn_Address3.Checked = additionsToLetter.Checked = 
                Memo_Box.Visible = turnMemoFormOnOff.Checked = false;
            excelFile.CreateAndRenewDropdownItems(names);
            excelFile.CreateAndRenewDropdownItems(corporations);
            excelFile.CreateAndRenewDropdownItems(senders);
            excelFile.CreateAndRenewDropdownItems(executors);

        }

        private void target_place_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var bookmarks = new string[] {"corporation_name", "corporation_address", "corporation_fax", "corporation_email" };
            var corporationInfo = excelFile.GetDataOfOneUnit(corporations, target_place.SelectedItem.ToString());
            //   var corporationInfo = LetterData.GetDataOfOneUnit(target_place.SelectedItem.ToString());
            if (!chk_address.Checked) corporationInfo[1] = " ";
            if (!chk_telfax.Checked) corporationInfo[2] = " ";
            if (!chk_email.Checked) corporationInfo[3] = " ";
            bookmarks = ChangeBookmarksIfButtonIsChosen(bookmarks);
            ChangeTextAtBookmark(bookmarks, corporationInfo);
            //Очищаем ФИО при смене организации
            var bookmarksOfPersonToDelete = ChangeBookmarksIfButtonIsChosen(new[] {"position", "shortname", "fullname"});
            ChangeTextAtBookmark(bookmarksOfPersonToDelete, new []{" ", " ", " "});

            ChangeCheckboxesInfoAndBack(true);
            excelFile.CreateAndRenewDropdownItems(names, target_place.SelectedItem.ToString());

        }

        private void target_person_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var bookmarks = new[] { "position", "shortname", "fullname"};
            //var personInfo = LetterData.GetPersonName(target_person.SelectedItem.ToString());
         //   var names = personalNameCl.CollectPersonNames;
            var personInfo = excelFile.GetDataOfOneUnit(names, target_person.SelectedItem.ToString());
            bookmarks = ChangeBookmarksIfButtonIsChosen(bookmarks);
            if (chk_isACopy.Checked) personInfo[2] = " ";
            ChangeTextAtBookmark(bookmarks, personInfo);
            ChangeCheckboxesInfoAndBack(true);
        }
        private void from_name_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var bookmarks = new[] { "from_position", "from_name"}; 
            //var personFromInfo = LetterData.GetFromNameAndStatus(from_name.SelectedItem.ToString());
         //   var senders = senderCl.CollectExecutors;
            var personFromInfo = excelFile.GetDataOfOneUnit(senders, from_name.SelectedItem.ToString());
            ChangeTextAtBookmark(bookmarks, personFromInfo);
        }

        private void creator_name_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var bookmarks = new string[] { "maker_name", "maker_number" };
            //var personFromInfo = LetterData.GetCreatorNameAndTel(creator_name.SelectedItem.ToString());
            var personFromInfo = excelFile.GetDataOfOneUnit(executors, creator_name.SelectedItem.ToString());
            ChangeTextAtBookmark(bookmarks, personFromInfo);
        }

        private void chk_address_Click(object sender, RibbonControlEventArgs e)
        {
          //  var corporationInfo = LetterData.GetCorporationName(target_place.SelectedItem.ToString());
            var corporationInfo = excelFile.GetDataOfOneUnit(corporations, target_place.SelectedItem.ToString());
            ChangePropertiesOfAdressOnline(chk_address, "corporation_address", corporationInfo[1]);
            ChangeCheckboxesInfoAndBack(true);
        }

        private void chk_telfax_Click(object sender, RibbonControlEventArgs e)
        {
            //var corporationInfo = LetterData.GetCorporationName(target_place.SelectedItem.ToString());
            var corporationInfo = excelFile.GetDataOfOneUnit(corporations, target_place.SelectedItem.ToString());
            ChangePropertiesOfAdressOnline(chk_telfax, "corporation_fax", corporationInfo[2]);
            ChangeCheckboxesInfoAndBack(true);
        }

        private void chk_email_Click(object sender, RibbonControlEventArgs e)
        {
            // var corporationInfo = LetterData.GetCorporationName(target_place.SelectedItem.ToString());
            var corporationInfo = excelFile.GetDataOfOneUnit(corporations, target_place.SelectedItem.ToString());
 
             ChangePropertiesOfAdressOnline(chk_email, "corporation_email", corporationInfo[3]);
            ChangeCheckboxesInfoAndBack(true);
        }

        private void turnLetterFormOnOff_Click(object sender, RibbonControlEventArgs e)
        {
            Letter_Box.Visible = turnLetterFormOnOff.Checked == true;
            Memo_Box.Visible = turnMemoFormOnOff.Checked = false;
            if (!turnMemoFormOnOff.Checked) return;
            excelFile.CreateAndRenewDropdownItems(names);
            excelFile.CreateAndRenewDropdownItems(corporations);
            excelFile.CreateAndRenewDropdownItems(senders);
            excelFile.CreateAndRenewDropdownItems(executors);
        }

        private void openLettersDirectory_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start("explorer.exe", @"\\SERVERRAID\z\Помойки\Помойка_Степана\Письма\");
        }

        private void additionsToLetter_Click(object sender, RibbonControlEventArgs e)
        {
            var bookmark = new[] { "additions" };
            var empty = new[] { " " };
            docBox_numberOfCopies.Visible =
                docBox_EditName.Visible = docBox_editLists.Visible = additionsToLetter.Checked;
            ChangeAdditionsToLetter();
            if (!additionsToLetter.Checked)
                ChangeTextAtBookmark(bookmark, empty);

        }

        private void btn_Address1_Click(object sender, RibbonControlEventArgs e)
        {
            btn_Address1.Checked = true;
            btn_Address2.Checked = btn_Address3.Checked = false;
            chk_isACopy.Visible = chk_isACopy.Checked = false;
            ChangeCheckboxesInfoAndBack(false);

        }

        private void btn_Address2_Click(object sender, RibbonControlEventArgs e)
        {
            var bookmarkArray = new[] { "additional_moreSpace_1", "additional_moreSpace_fullname_1" };
            var newText = new string[] { "\n\n", "\n"};
            btn_Address1.Checked = btn_Address3.Checked = false;
            btn_Address2.Checked = chk_isACopy.Visible = true;
            ChangeTextAtBookmark(bookmarkArray,newText);
            ChangeCheckboxesInfoAndBack(false);
        }

        private void btn_Address3_Click(object sender, RibbonControlEventArgs e)
        {
            var bookmarkArray = new[] { "additional_moreSpace_2", "additional_moreSpace_fullname_2" };
            var newText = new string[] { "\n\n", "\n" };
            btn_Address1.Checked = btn_Address2.Checked = false;
            btn_Address3.Checked = chk_isACopy.Visible = true;
            ChangeTextAtBookmark(bookmarkArray, newText);
            ChangeCheckboxesInfoAndBack(false);
        }

        private void chk_isACopy_Click(object sender, RibbonControlEventArgs e)
        {
            if (chk_isACopy.Checked)
            {
                var bookmarks = new[] {"isACopy", "fullname" , "moreSpace_fullname"};
                var newText = new[] {"Копия:", " ", " "};
                ChangeBookmarksIfButtonIsChosen(bookmarks);
                ChangeCheckboxesInfoAndBack(true); 
                ChangeTextAtBookmark(bookmarks,newText);
            }
            else
            {
                var bookmarks = new[] { "isACopy", "fullname", "moreSpace_fullname" };
                ChangeBookmarksIfButtonIsChosen(bookmarks);
                //var personInfo = LetterData.GetPersonName(target_person.SelectedItem.ToString());
                var personInfo = excelFile.GetDataOfOneUnit(names, target_person.SelectedItem.ToString());
                var newText = new[] {" ", personInfo[2], "\n"};
                ChangeCheckboxesInfoAndBack(true);
                ChangeTextAtBookmark(bookmarks, newText);
            }
        }
        private void docBox_editLists_TextChanged(object sender, RibbonControlEventArgs e)
        {
            ChangeAdditionsToLetter();
        }

        private void docBox_EditName_TextChanged(object sender, RibbonControlEventArgs e)
        {
            ChangeAdditionsToLetter();
        }

        private void docBox_numberOfCopies_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            ChangeAdditionsToLetter();
        }
        private void templates_dropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            ChangeTextAtSelection();
        }

        #endregion

        #region Функциональная часть

        private void ChangeTextAtBookmark(string[] bookmarks, string[] info)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            for (int i = 0; i < bookmarks.Length; i++)
            {
                try
                {
                    Word.Bookmark bm = document.Bookmarks[bookmarks[i]];
                    Word.Range range = bm.Range;
                    range.Text = info[i];
                    document.Bookmarks.Add(bookmarks[i], range);
                }
                catch 
                {
                    return;
                }
            }
        }
        
        private string[] ChangeBookmarksIfButtonIsChosen(string[] bookmarks)
        {
            
            if (btn_Address2.Checked)
            {
                for (int i = 0; i < bookmarks.Length; i++)
                {
                    bookmarks[i] = "additional_" + bookmarks[i] + "_1";
                }
            }
            if (btn_Address3.Checked)
            {
                for (int i = 0; i < bookmarks.Length; i++)
                {
                    bookmarks[i] = "additional_" + bookmarks[i] + "_2";
                }
            }

            return bookmarks;
        }

        private void ChangePropertiesOfAdressOnline(RibbonCheckBox checkBox,
            string bookmarkName, string newInformation)
        {
            var bookmark = new string[] {bookmarkName};
            var infoToChange = new string[] {newInformation};
            var emptySpace = new string[1];
            bookmark = ChangeBookmarksIfButtonIsChosen(bookmark);
             ChangeTextAtBookmark(bookmark, checkBox.Checked ? infoToChange : emptySpace);
        }

        private void ChangeCheckboxesInfoAndBack(bool write)
        {
            if (write)
            {
                if (btn_Address1.Checked)
                    SaveAddressStatusInfo(1);
                if (btn_Address2.Checked)
                    SaveAddressStatusInfo(2);
                if (btn_Address3.Checked)
                    SaveAddressStatusInfo(3);
            }
            else
            {
                if (btn_Address1.Checked)
                    LoadAddressStatusInfo(1);
                if (btn_Address2.Checked)
                    LoadAddressStatusInfo(2);
                if (btn_Address3.Checked)
                    LoadAddressStatusInfo(3);
            }
        }
        private void SaveAddressStatusInfo(int id)
        {
            switch (id)
            {
                case 1:
                    firstAddressInfo.AddressCheckBox = chk_address.Checked;
                    firstAddressInfo.EmailCheckBox = chk_email.Checked;
                    firstAddressInfo.TelFaxCheckBox = chk_telfax.Checked;
                    firstAddressInfo.IsACopyCheckBox = false;
                    firstAddressInfo.SelectedCorporation = target_place.SelectedItem;
                    firstAddressInfo.SelectedPersonName = target_person.SelectedItem;
                    break;
                case 2:
                    secondAddressInfo.AddressCheckBox = chk_address.Checked;
                    secondAddressInfo.EmailCheckBox = chk_email.Checked;
                    secondAddressInfo.TelFaxCheckBox = chk_telfax.Checked;
                    secondAddressInfo.IsACopyCheckBox = chk_isACopy.Checked;
                    secondAddressInfo.SelectedCorporation = target_place.SelectedItem;
                    secondAddressInfo.SelectedPersonName = target_person.SelectedItem;
                    break;
                case 3:
                    thirdAddressInfo.AddressCheckBox = chk_address.Checked;
                    thirdAddressInfo.EmailCheckBox = chk_email.Checked;
                    thirdAddressInfo.TelFaxCheckBox = chk_telfax.Checked;
                    thirdAddressInfo.IsACopyCheckBox = chk_isACopy.Checked;
                    thirdAddressInfo.SelectedCorporation = target_place.SelectedItem;
                    thirdAddressInfo.SelectedPersonName = target_person.SelectedItem;
                    break;
            }
        }

        private void LoadAddressStatusInfo(int id)
        {
            switch (id)
            {
                case 1:
                    chk_address.Checked = firstAddressInfo.AddressCheckBox;
                    chk_telfax.Checked = firstAddressInfo.TelFaxCheckBox;
                    chk_email.Checked = firstAddressInfo.EmailCheckBox;
                    //   target_person.SelectedItem = firstAddressInfo.SelectedPersonName; // TODO: Реализовать позже
                    target_place.SelectedItem = firstAddressInfo.SelectedCorporation;
                    //excelFile.CreateAndRenewDropdownItems(personalNameCl.CollectPersonNames, target_place.SelectedItem.Label);
                    excelFile.CreateAndRenewDropdownItems(names, target_place.SelectedItem.Label);
                    break;
                case 2:
                    chk_address.Checked = secondAddressInfo.AddressCheckBox;
                    chk_telfax.Checked = secondAddressInfo.TelFaxCheckBox;
                    chk_email.Checked = secondAddressInfo.EmailCheckBox;
                    //    target_person.SelectedItem = secondAddressInfo.SelectedPersonName;
                    target_place.SelectedItem = secondAddressInfo.SelectedCorporation;
                    chk_isACopy.Checked = secondAddressInfo.IsACopyCheckBox;
                    excelFile.CreateAndRenewDropdownItems(names, target_place.SelectedItem.Label);
                    break;
                case 3:
                    chk_address.Checked = thirdAddressInfo.AddressCheckBox;
                    chk_telfax.Checked = thirdAddressInfo.TelFaxCheckBox;
                    chk_email.Checked = thirdAddressInfo.EmailCheckBox;
                    //    target_person.SelectedItem = thirdAddressInfo.SelectedPersonName;
                    target_place.SelectedItem = thirdAddressInfo.SelectedCorporation;
                    chk_isACopy.Checked = thirdAddressInfo.IsACopyCheckBox;
                    excelFile.CreateAndRenewDropdownItems(names, target_place.SelectedItem.Label);
                    break;
                default:
                    chk_address.Checked = true;
                    chk_telfax.Checked = true;
                    chk_email.Checked = true;
                    chk_isACopy.Checked = false;
                    target_person.SelectedItem = target_person.SelectedItem;
                    target_place.SelectedItem = target_place.SelectedItem;
                    break;
            }
        }

        private void ChangeAdditionsToLetter()
        {
            if (!int.TryParse(docBox_editLists.Text, out var numberOfLists))
            {
                docBox_editLists.Text = "1";
                numberOfLists = 0;
            }
            var text = docBox_EditName.Text;
            var numberOfCopies = int.Parse(docBox_numberOfCopies.SelectedItem.Label);

            var totalLists = numberOfLists * numberOfCopies;
            if (text == " " || text == "")
                text = "Приложение по тексту";
            var additions = new[] { "Приложение: " + text + " на " + numberOfLists + " л. в " + numberOfCopies + " экз., всего на "+ totalLists+" л." };
            var bookmark = new[] { "additions" };
            if (numberOfCopies == 1)
                additions = new[] { "Приложение: " + text + " на " + numberOfLists + " л."};
            ChangeTextAtBookmark(bookmark, additions);
        }

        private void ChangeTextAtSelection() // from MSDN
        {
            var application = Globals.ThisAddIn.Application;
            var currentSelection = application.Selection;
            // Store the user's current Overtype selection
            bool userOvertype = application.Options.Overtype;

            // Make sure Overtype is turned off.
            if (application.Options.Overtype)
            {
                application.Options.Overtype = false;
            }

            // Test to see if selection is an insertion point.
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionIP)
            {
                currentSelection.TypeText(excelFile.PreparingTextToInsert(templates, templates_dropDown.SelectedItem.ToString()));
            }
            else
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionNormal)
            {
                // Inserting before text block;
                // Move to start of selection.
                if (application.Options.ReplaceSelection)
                {
                    object direction = Word.WdCollapseDirection.wdCollapseStart;
                    currentSelection.Collapse(ref direction);
                }
                currentSelection.TypeText(excelFile.PreparingTextToInsert(templates, templates_dropDown.SelectedItem.ToString()));
                currentSelection.TypeParagraph();
            }
            // Restore the user's Overtype selection
            application.Options.Overtype = userOvertype;
        }


        #endregion

        #region Поля информации о письме

        public class LetterInfo
        { 
            public bool AddressCheckBox { get; set; }
            public bool TelFaxCheckBox { get; set; }
            public bool EmailCheckBox { get; set; }
            public bool IsACopyCheckBox { get; set; }
            public RibbonDropDownItem SelectedCorporation { get; set; }
            public RibbonDropDownItem SelectedPersonName { get; set; }
            
        }

        #endregion

        #region Кнопки управления служебной запиской

        private void turnMemoFormOnOff_Click(object sender, RibbonControlEventArgs e)
        {
            turnLetterFormOnOff.Checked = Letter_Box.Visible = false;
            Memo_Box.Visible = turnMemoFormOnOff.Checked == true;
            btnMemoTemp.Enabled = true;

        }

        private void btnMemoTemp_Click(object sender, RibbonControlEventArgs e)
        {
            object missing = System.Type.Missing;
            try
            {
                Globals.ThisAddIn.Application.Documents.Add(
                    @"\\SERVERRAID\z\Помойки\Помойка_Степана\Programms\orion_forms\memo_form.docx", ref missing,
                    ref missing, ref missing);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Ошибка при обращении к файлу формы:\n" + exception.Message);
            }
            Memo_Box.Visible = turnMemoFormOnOff.Checked = true;
            Letter_Box.Visible = turnLetterFormOnOff.Checked = false;
        }

        private void memoTargetPerson_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var bookmarks = new string[] { "memo_target_person" };
            var personInfo = MemoData.GetPersonName(memoTargetPerson.SelectedItem.ToString());
            ChangeTextAtBookmark(bookmarks, personInfo);
        }

        private void memoFromPerson_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var bookmarks = new string[] { "memo_from_person" };
            var personInfo = MemoData.GetFromName(memoFromPerson.SelectedItem.ToString());
            ChangeTextAtBookmark(bookmarks, personInfo);
        }

        private void btn_InsertDate_Click(object sender, RibbonControlEventArgs e)
        {
            var date = DateTime.Now;
            var dateToInsert = new string[] {date.ToString("d")};
            var bookmarks = new string[] {"memo_enterdate"};
            var empty = new string[] {" "};
            ChangeTextAtBookmark(bookmarks, btn_InsertDate.Checked ? dateToInsert : empty);
        }


        #endregion

    }
}
