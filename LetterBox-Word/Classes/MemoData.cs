using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LetterBox_Word
{
    [Obsolete("This class will be changed to ExcelDBWork equal")]
    public class MemoData
    {
        public static string[] GetPersonName(string shortname)
        {
            var result = new string[1];
            var orion = "\nАО «НПО «Орион»";

            switch (shortname)
            {
                case "Кузнецов С.А.":
                    result[0] = "Временному генеральному директору" + orion + "\nКузнецову С.А.";
                    break;
                case "Федоров А.Г.":
                    result[0] = "Заместителю генерального директора по безопасности и режиму" + orion +
                                "\nФедорову А.Г.";
                    break;
                case "Конькова И.Г.":
                    result[0] = "Зам. ген. директора по экономическому развитию и управлению финансами" + orion +
                                "\nКоньковой И.Г.";
                    break;
                case "Бурлаков И.Д.":
                    result[0] = "Заместителю генерального директора по инновациям и науке" + orion + 
                                "\nБурлакову И.Д.";
                    break;
                case "Гринченко Л.Я.":
                    result[0] = "Начальнику УИТ" + "\nГринченко Л.Я.";
                    break;
                case "Еникеев О.И.":
                    result[0] = "Начальнику УМР" + "\nЕникееву О.И.";
                    break;
                case "Ефимова З.Н.":
                    result[0] = "Заместителю начальника НТК" + "\nЕфимовой З.Н.";
                    break;
                case "Болтарь К.О.":
                    result[0] = "Начальнику НТК" + "\nБолтарю К.О.";
                    break;
                case "Кульчицкий Н.А.":
                    result[0] = "Зам. начальника УПП и СР" + "\nКульчицкому Н.А.";
                    break;
                case "Бучинская Н.В.":
                    result[0] = "Начальнику отдела кадров" + "\nБучинской Н.В.";
                    break;
                case "Янкина М.В.":
                    result[0] = "Начальнику финансового управления" + "\nЯнкиной М.В.";
                    break;
                case "Дворак А.И.":
                    result[0] = "Главному бухгалтеру" + "\nДворак А.И.";
                    break;
                case "Полесский А.В.":
                    result[0] = "Главному метрологу" + "\nПолесскому А.В.";
                    break;
                default:
                    result[0] = " ";
                    break;
            }


            return result;
        }

        public static string[] GetFromName(string shortname)
        {
            var result = new string[1];
            switch (shortname)
            {
                case "Дражников Б.Н.":
                    result[0] = "от начальника НТЦ №2" + "\nДражникова Б.Н.";
                    break;
                case "Бычковский Я.С.":
                    result[0] = "от заместителя начальника научно-технического центра №2" + "\nБычковского Я.С.";
                    break;
                default:
                    result[0] = " ";

                    break;
            }

            return result;
        }

    }
}