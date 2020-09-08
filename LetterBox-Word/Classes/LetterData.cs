using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LetterBox_Word
{
    [Obsolete("This class will not be used anymore")]
    public class LetterData
    {
        public static string[] GetCorporationName(string shortname)
        {
            var result = new string[4];
            var fax = "Факс: ";
            var tel = "Тел: ";
            var telfax = "Тел/Факс: ";
            var email = "E-mail: ";
            var next = "\n";
            switch (shortname)
            {
                case "Комета Москва":
                    result[0] = "АО «Корпорация «Комета»";
                    result[1] = "Велозаводская ул., д.5,\n г. Москва, 115280";
                    result[2] = telfax + "+7 (495) 674-08-46";
                    result[3] = email + "info@corpkometa.ru";
                    break;
                case "НПЦ ОЭКН Комета Спб":
                    result[0] = "АО «Корпорация «Комета» –\n«НПЦ ОЭКН»";
                    result[1] = "ул. Шателена, д. 7\nг. Санкт - Петербург, 194021";
                    result[2] = telfax + "+7 (812) 331-61-00";
                    result[3] = email + "kometa@eoss.ru";
                    break;
                case "Субмикрон":
                    result[0] = "AО «НИИ «Субмикрон»";
                    result[1] = "Георгиевский пр-т, д. 5, стр. 2\n г. Москва, Зеленоград, 124460";
                    result[2] = fax + "+7 (499) 731-27-53";
                    result[3] = email + "submicron@se.zgrad.ru";
                    break;
                case "ОКТБ Омега":
                    result[0] = "АО «ОКТБ «Омега»";
                    result[1] = "ул. Саши Устинова, д. 1\n г. Великий Новгород, 173003";
                    result[2] = tel + "+7 (8162) 62-64-02\n" + fax + "+7 (8162) 62-67-85";
                    result[3] = email + "omega@oktb-omega.ru";
                    break;
                case "ИАиЭ СО РАН":
                    result[0] = "Института автоматики и\nэлектрометрии Сибирского отделения\n Российской академии наук\n(ИАиЭ СО РАН)";
                    result[1] = "проспект Академика Коптюга, д. 1\nг. Новосибирск, 630090";
                    result[2] = fax + "+7 (383) 330-88-78";
                    result[3] = email + "iae@iae.nsk.su";
                    break;
                case "НПО Лавочкина":
                    result[0] = "АО «НПО Лавочкина»";
                    result[1] = "ул. Ленинградская, д.24\nМосковская область, г. Химки, 141400";
                    result[2] = tel + "+7 (495) 573-56-75, +7 (495) 575-55-11\n" + fax + "+7 (495) 573-35-95";
                    result[3] = email + "npol@laspace.ru";
                    break;
                case "НИИМЭ":
                    result[0] = "АО «НИИМЭ»";
                    result[1] = "ул. Академика Валиева, д. 6, стр. 1\n г. Москва, Зеленоград, 124460";
                    result[2] = tel + "+7 (495) 229-72-99\n" + fax + "+7 (495) 229-77-73";
                    result[3] = email + "niime@niime.ru";
                    break;
                case "333 ВП МО - Комета Мск":
                    result[0] = "333 ВП МО РФ";
                    result[1] = "Велозаводская ул., д.5,\nМосква, 115280";
                    result[2] = telfax + "+7 (495) 674-08-46";
                    result[3] = email + "info@corpkometa.ru";
                    break;
                case "384 ВП МО - Комета Спб":
                    result[0] = "384 ВП МО РФ";
                    result[1] = "ул. Шателена, д. 7\nг.Санкт - Петербург, 194021";
                    result[2] = telfax + "+7 (812) 331-61-00";
                    result[3] = email + "kometa@eoss.ru";
                    break;
                case "628 ВП МО - ИАиЭ СО РАН":
                    result[0] = "628 ВП МО РФ";
                    result[1] = "ул. Дуси Ковальчук, д. 276\nг. Новосибирск, 630075\nАО «НПП «Восток»";
                    result[2] = tel + "+7 (383) 225-63-49";
                    result[3] = " ";
                    break;
                case "263 ВП МО - Омега":
                    result[0] = "263 ВП МО РФ";
                    result[1] = "ул. Саши Устинова, д. 1\n г. Великий Новгород, 173003";
                    result[2] = tel + "+7 (8162) 62-64-02\n" + fax + "+7 (8162) 62-67-85";
                    result[3] = email + "omega@oktb-omega.ru";
                    break;
                case "3960 ВП МО - Субмикрон":
                    result[0] = "3960 ВП МО РФ";
                    result[1] = "Георгиевский пр-т, д. 5, стр. 2\n г. Москва, Зеленоград, 124460";
                    result[2] = fax + "+7 (499) 731-27-53";
                    result[3] = email + "submicron@se.zgrad.ru";
                    break;
                case "4116 ВП МО - Лавочкина":
                    result[0] = "4116 ВП МО РФ";
                    result[1] = "ул. Ленинградская, д.24\nМосковская область, г. Химки, 141400";
                    result[2] = tel + "+7 (495) 573-56-75, +7 (495) 575-55-11\n" + fax + "+7 (495) 573-35-95";
                    result[3] = email + "npol@laspace.ru";
                    break;
                case "514 ВП МО - НИИМЭ":
                    result[0] = "514 ВП МО РФ";
                    result[1] = "ул. Академика Валиева, д. 6, стр. 1\n г. Москва, Зеленоград, 124460";
                    result[2] = tel + "+7 (495) 229-72-99\n" + fax + "+7 (495) 229-77-73";
                    result[3] = email + "niime@niime.ru";
                    break;
                case "524 ВП МО - Орион":
                    result[0] = "524 ВП МО РФ";
                    result[1] = "ул. Косинская д. 9\n г. Москва, 111123";
                    result[2] = tel + "+7 (495) 672-20-17\n";
                    result[3] = email + "orion@orion-ir.ru";
                    break;
                default:
                    result[0] = " ";
                    result[1] = " ";
                    result[2] = " ";
                    result[3] = " ";
                    break;
            }

            for (int i = 0; i < 4; i++)
            {
                result[i] = next + result[i];
            }
            return result;
        }

        public static string[] GetPersonName(string shortname)
        {
            var result = new string[3];
            var goodOne = "Уважаемый ";
            var nextParagraph = "\n";
            var exclamation = "!";
            switch (shortname)
            {
                case "Мисник - Комета Мск":
                    result[0] = "Генеральному директору –\n генеральному конструктору";
                    result[1] = "Миснику В.П.";
                    result[2] = goodOne + "Виктор Порфирьевич!";
                    break;
                case "Бодин - Комета Мск":
                    result[0] = "Главному инженеру";
                    result[1] = "Бодину В.В.";
                    result[2] = goodOne + "Вадим Витальевич!";
                    break;
                case "Захаров - Комета Мск":
                    result[0] = "Заместителю генерального директора,\n заместителю генерального\n конструктора ЕКС";
                    result[1] = "Захарову А.А.";
                    result[2] = goodOne + "Андрей Анатольевич!";
                    break;
                case "Погребский - НПЦОЭКН":
                    result[0] = "Директору филиала";
                    result[1] = "Погребскому Н.А.";
                    result[2] = goodOne + "Николай Аркадьевич!";
                    break;
                case "Парпин - НПЦОЭКН":
                    result[0] = "Заместителю директора филиала \n по разработкам";
                    result[1] = "Парпину М.А.";
                    result[2] = goodOne + "Михаил Анатольевич!";
                    break;
                case "Орлов - Субмикрон":
                    result[0] = "Генеральному директору";
                    result[1] = "Орлову А.В.";
                    result[2] = goodOne + "Алексей Валерьевич!";
                    break;
                case "Гришин - Субмикрон":
                    result[0] = "Первому заместителю \n генерального директора \n Главному конструктору";
                    result[1] = "Гришину В.Ю.";
                    result[2] = goodOne + "Вячеслав Юрьевич!";
                    break;
                case "Комаревцев - Омега":
                    result[0] = "Генеральному директору";
                    result[1] = "Комаревцеву Д.В.";
                    result[2] = goodOne + "Дмитрий Владимирович!";
                    break;
                case "Корольков - ИАиЭ СО РАН":
                    result[0] = "Заместителю директора по научной работе";
                    result[1] = "Королькову В.П.";
                    result[2] = goodOne + "Виктор Павлович!";
                    break;
                case "Бабин - ИАиЭ СО РАН":
                    result[0] = "Директору";
                    result[1] = "Бабину С.А.";
                    result[2] = goodOne + "Сергей Алексеевич!";
                    break;
                case "Поляков - Лавочкина":
                    result[0] = "Заместителю генерального конструктора по механическим системам";
                    result[1] = "Полякову А.А.";
                    result[2] = goodOne + "Алексей Александрович!";
                    break;
                case "Красников - НИИМЭ":
                    result[0] = "Генеральному директору";
                    result[1] = "Красникову Г.Я.";
                    result[2] = goodOne + "Геннадий Яковлевич!";
                    break;
                case "Шелепин - НИИМЭ":
                    result[0] = "Первому заместителю генерального директора";
                    result[1] = "Шелепину Н.А.";
                    result[2] = goodOne + "Николай Алексеевич!";
                    break;
                case "Байкин - 333 ВП МО":
                    result[0] = "Начальнику";
                    result[1] = "Байкину Е.А.";
                    result[2] = goodOne + "Егор Александрович!";
                    break;
                case "Губань - 384 ВП МО":
                    result[0] = "Начальнику";
                    result[1] = "Губань Д.К.";
                    result[2] = goodOne + "Денис Константинович";
                    break;
                case "Аксенов - 263 ВП МО":
                    result[0] = "Начальнику";
                    result[1] = "Аксенову С.С.";
                    result[2] = goodOne + "Станислав Сергеевич";
                    break;
                case "Тимаков - 628 ВП МО":
                    result[0] = "Начальнику";
                    result[1] = "Тимакову О.Э.";
                    result[2] = goodOne + "Олег Эдуардович!";
                    break;
                case "Широкорад - 3960 ВП МО":
                    result[0] = "Начальнику";
                    result[1] = "Широкораду А.Е.";
                    result[2] = goodOne + "Александр Евгеньевич!";
                    break;
                case "Байкин - 4116 ВП МО":
                    result[0] = "Начальнику";
                    result[1] = "Байкину В.В.";
                    result[2] = goodOne + "Виталий Владимирович!";
                    break;
                case "Швид - 514 ВП МО":
                    result[0] = "Начальнику";
                    result[1] = "Швиду И.А.";
                    result[2] = goodOne + "Игорь Александрович!";
                    break;
                case "Бубнов - 524 ВП МО":
                    result[0] = "Начальнику отдела";
                    result[1] = "Бубнову С.В.";
                    result[2] = goodOne + "Сергей Владимирович!";
                    break;
                case "Планида - 524 ВП МО":
                    result[0] = "Заместителю начальника отдела";
                    result[1] = "Планиде С.С.";
                    result[2] = goodOne + "Сергей Сергеевич!";
                    break;
                default:
                    result[0] = " ";
                    result[1] = " ";
                    result[2] = " ";
                    break;
            }

            result[1] = nextParagraph + result[1] + nextParagraph;
            return result;
        }

        public static string[] GetFromNameAndStatus(string shortname)
        {
            var result = new string[2];
            var regards = "С уважением,\n\n";
            switch (shortname)
            { 
                case "Кузнецов С.А.":
                    result[0] = regards + "Временный генеральный директор";
                    result[1] = "\nС.А. Кузнецов";
                    break;
                case "Бурлаков И.Д.":
                    result[0] = regards + "Заместитель генерального\nдиректора по инновациям и науке";
                    result[1] = "\n\nИ.Д. Бурлаков";
                    break;
                case "Конькова И.Г.":
                    result[0] =
                        regards +
                        "Заместитель генерального директора\nпо экономическому развитию и\nуправлению финансами";
                    result[1] = "\n\n\nИ.Г. Конькова";
                    break;
                case "Михайличенко С.А. (за Бурлакова)":
                    result[0] = regards + "ВрИО заместителя генерального\nдиректора по инновациям и науке";
                    result[1] = "\n\nС.А. Михайличенко";
                    break;
                default:
                    result[0] = " ";
                    result[1] = " ";
                    break;
            }

            return result;
        }

        public static string[] GetCreatorNameAndTel(string shortname)
        {
            var result = new string[2];
            var numberNTC = "+7 (499) 374-47-60";
            var isp = " Исп. ";
            var tel = "Тел. ";
            switch (shortname)
            {
                case "Бычковский Я.С.":
                    result[0] = isp + "Бычковский Я.С.\n";
                    result[1] = tel + "+7 (909) 986-77-13";
                    break;
                case "Филиппов С.О.":
                    result[0] = isp + "Филиппов С.О.\n";
                    result[1] = tel + numberNTC;
                    break;
                case "Филиппова И.С.":
                    result[0] = isp + "Филиппова И.С.\n";
                    result[1] = tel + numberNTC;
                    break;
                case "Леонтьев Е.В.":
                    result[0] = isp + "Леонтьев Е.В.\n";
                    result[1] = tel + numberNTC;
                    break;
                case "Кондюшин И.С.":
                    result[0] = isp + "Кондюшин И.С.\n";
                    result[1] = tel + numberNTC;
                    break;
                case "Крехова Е.Ю.":
                    result[0] = isp + "Крехова Е.Ю.\n";
                    result[1] = tel + numberNTC;
                    break;
                default:
                    result[0] = " \n";
                    result[1] = " ";
                    break;

            }

            return result;
        }
    }

}