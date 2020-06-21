using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataSetExcel
{
    //class myColumn
    //{
    //    public int value;
    //    public string columnName;
    //}

    class Program
    {
        static DateTime DTMorning = new DateTime();
        static DateTime DTDay = new DateTime();
        static DateTime DTEvening = new DateTime();

        static void Main(string[] args)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.ScreenUpdating = false;
            ExcelApp.Visible = false;
            ExcelApp.Interactive = false;
            ExcelApp.IgnoreRemoteRequests = true;

            DTMorning = DTMorning.AddHours(6);
            DTDay = DTDay.AddHours(12);
            DTEvening = DTEvening.AddHours(18);
            //var Month = Enumerable.Range(1, 12).Select(i => new { I = i, M = DateTimeFormatInfo.CurrentInfo.GetMonthName(i) });

            // Массив праздников в ноябре 2018 и 2019
            DateTime[] PartyMassive = new DateTime[] {
                DateTime.Parse("04.11.2018"),
                DateTime.Parse("05.11.2018"),
                DateTime.Parse("01.01.2019"),
                DateTime.Parse("02.01.2019"),
                DateTime.Parse("03.01.2019"),
                DateTime.Parse("04.01.2019"),
                DateTime.Parse("05.01.2019"),
                DateTime.Parse("06.01.2019"),
                DateTime.Parse("07.01.2019"),
                DateTime.Parse("08.01.2019"),
                DateTime.Parse("23.02.2019"),
                DateTime.Parse("03.05.2019"),
                DateTime.Parse("01.05.2019"),
                DateTime.Parse("02.05.2019"),
                DateTime.Parse("03.05.2019"),
                DateTime.Parse("09.05.2019"),
                DateTime.Parse("10.05.2019"),
                DateTime.Parse("12.06.2019"),
                DateTime.Parse("04.11.2019"),
                DateTime.Parse("31.12.2019"),
            };

            List<string> dayOfWeekMassive = new List<string>(new string[] { "понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье", });

            string sheetName = @"Data";                 // Название листа откуда берётся информация
            string sheetResultName = @"DataResult";     // Название листа откуда куда поместятся преобразованные данные
            string sheetExcplanation = @"Excplanation"; // Название листа где будет обозначение данных

            Console.WriteLine("Название листа, откуда берётся информация должно быть: " + sheetName);
            string fName = @"C:\Users\35498\source\repos\DataSetExcel\Neuro\NewData\DataT.xlsx"; // Файл Excel, с которым производится работа
            Console.WriteLine("Выбрать файл?: " + fName + " ? (n for no)");
            string str = Console.ReadLine();
            if (str == "n")
            {
                Console.WriteLine("Введите файл");
                fName = Console.ReadLine();
            }

            if (str.Contains(@":\")) { fName = str; }

            Console.WriteLine("выполнение..");

            // Название колонок, находящееся в первой строки листа Excel, откуда будут поступать данные
            string COLUMN_DATA_NAME_FROM = "Дата";
            string COLUMN_TIME_NAME_FROM = "Время";
            string COLUMN_TYPEDTP_NAME_FROM = "Вид ДТП";
            string COLUMN_ROAD_NAME_FROM = "Дорога";
            string COLUMN_KILOMETR_NAME_FROM = "Километр";
            string COLUMN_METR_NAME_FROM = "Метр";
            string COLUMN_NDU_NAME_FROM = "НДУ";
            string COLUMN_FACTORS_NAME_FROM = "Факторы, влияющие на режим движения";
            string COLUMN_STATUSROAD_NAME_FROM = "Состояние проезжей части";
            string COLUMN_STATUSWEATHER_NAME_FROM = "Состояние погоды";
            string COLUMN_LIGHT_NAME_FROM = "Освещение";
            string COLUMN_POINT_NAME_FROM = "Является местом концентрации ДТП";
            string COLUMN_BAD_NAME_FROM = "Непосредственные нарушения ПДД";

            // Название колонок, которые будут в результирующем листе
            string COLUMN_DAY_NAME = "День";
            string COLUMN_MONTH_NAME = "Месяц";
            string COLUMN_WEEK_NAME = "День недели";
            string COLUMN_PARTY_NAME = "Праздник";
            string COLUMN_TIMEOFDAY_NAME = "Время";
            string COLUMN_TYPEDTP_NAME = "Вид ДТП";
            string COLUMN_ROAD_NAME = "Дорога";
            string COLUMN_KILOMETR_NAME = "Километр";
            string COLUMN_METR_NAME = "Метр";
            string COLUMN_NDU_NAME = "НДУ";
            string COLUMN_FACTORS_NAME = "Факторы, влияющие на режим движения";
            string COLUMN_STATUSROAD_NAME = "Состояние проезжей части";
            string COLUMN_STATUSWEATHER_NAME = "Состояние погоды";
            string COLUMN_LIGHT_NAME = "Освещение";
            string COLUMN_POINT_NAME = "Является местом концентрации ДТП";
            string COLUMN_BAD_NAME = "Непосредственные нарушения ПДД";

            // Название колонок, обозначающих сопоставленное ячейке число
            string SUFFIX_IN_EXPLANATION = "Число";

            var wb = ExcelApp.Workbooks.Open(fName);

            try
            {
                var sheet = (Excel.Worksheet)wb.Worksheets[sheetName];
                Excel.Worksheet resultSheet;
                Excel.Worksheet ExcplanationSheet;

                gg1:
                try
                {
                    resultSheet = (Excel.Worksheet)wb.Worksheets[sheetResultName];
                    sheetResultName += "1";
                    goto gg1;
                }
                catch { }

                gg2:
                try
                {
                    ExcplanationSheet = (Excel.Worksheet)wb.Worksheets[sheetExcplanation];
                    sheetExcplanation += "1";
                    goto gg2;
                }
                catch { }

                ((Excel.Worksheet)wb.Worksheets.Add()).Name = sheetResultName;
                resultSheet = (Excel.Worksheet)wb.Worksheets[sheetResultName];
                ((Excel.Worksheet)wb.Worksheets.Add()).Name = sheetExcplanation;
                ExcplanationSheet = (Excel.Worksheet)wb.Worksheets[sheetExcplanation];

                int columnFromData = GetColumnIndex(sheet, COLUMN_DATA_NAME_FROM);
                int columnFromTime = GetColumnIndex(sheet, COLUMN_TIME_NAME_FROM);
                int columnFromTypeDTP = GetColumnIndex(sheet, COLUMN_TYPEDTP_NAME_FROM);
                int columnFromRoad = GetColumnIndex(sheet, COLUMN_ROAD_NAME_FROM);
                int columnFromKilometr = GetColumnIndex(sheet, COLUMN_KILOMETR_NAME_FROM);
                int columnFromMetr = GetColumnIndex(sheet, COLUMN_METR_NAME_FROM);
                int columnFromNDU = GetColumnIndex(sheet, COLUMN_NDU_NAME_FROM);
                int columnFromFactors = GetColumnIndex(sheet, COLUMN_FACTORS_NAME_FROM);
                int columnFromStatRoad = GetColumnIndex(sheet, COLUMN_STATUSROAD_NAME_FROM);
                int columnFromStatWeather = GetColumnIndex(sheet, COLUMN_STATUSWEATHER_NAME_FROM);
                int columnFromLight = GetColumnIndex(sheet, COLUMN_LIGHT_NAME_FROM);
                int columnFromPoint = GetColumnIndex(sheet, COLUMN_POINT_NAME_FROM);
                int columnFromBAD = GetColumnIndex(sheet, COLUMN_BAD_NAME_FROM);

                StringDifferentValueHandler sdvhTypeDTP = new StringDifferentValueHandler();
                StringDifferentValueHandler sdvhRoad = new StringDifferentValueHandler();
                StringDifferentValueHandler sdvhNDU = new StringDifferentValueHandler();            //
                StringDifferentValueHandler sdvhFactor = new StringDifferentValueHandler();         //
                StringDifferentValueHandler sdvhStatRoad = new StringDifferentValueHandler();
                StringDifferentValueHandler sdvhStatWeather = new StringDifferentValueHandler();
                StringDifferentValueHandler sdvhLight = new StringDifferentValueHandler();
                StringDifferentValueHandler sdvhBAD = new StringDifferentValueHandler();            //
                StringDifferentValueHandler sdvhWeekDays = new StringDifferentValueHandler();
                sdvhWeekDays.Add(dayOfWeekMassive.ToArray());

                ///////////----------------------------////////////////////
                int i = 2;
                for (i = 2; sheet.Cells[i, 1].Value != null; i++)
                {
                    sdvhNDU.Add(ParseString(sheet.Cells[i, columnFromNDU].Text).ToArray());

                    sdvhFactor.Add(ParseString(sheet.Cells[i, columnFromFactors].Text).ToArray());

                    sdvhBAD.Add(ParseString(sheet.Cells[i, columnFromBAD].Text).ToArray());
                }

                int columnIndex = 1;
                resultSheet.Cells[1, columnIndex] = COLUMN_DAY_NAME;
                int res_COLUMN_DAY = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_MONTH_NAME;
                int res_COLUMN_MONTH = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_WEEK_NAME;
                int res_COLUMN_WEEK = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_PARTY_NAME;
                int res_COLUMN_PARTY = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_TIMEOFDAY_NAME;
                int res_TIMEOFDAY_PARTY = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_TYPEDTP_NAME;
                int res_TYPEDTP = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_ROAD_NAME;
                int res_ROAD = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_KILOMETR_NAME;
                int res_KILOMETR = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_METR_NAME;
                int res_METR = columnIndex;

                //**//
                //columnIndex++;
                int resBegin_NDU = columnIndex + 1;
                for (int k = 0; k < sdvhNDU.Values.Count; k++)
                    resultSheet.Cells[1, resBegin_NDU + k] = sdvhNDU.Values[k];
                //**//
                columnIndex += sdvhNDU.Values.Count + 1;
                int resBegin_FACTOR = columnIndex;
                for (int k = 0; k < sdvhFactor.Values.Count; k++)
                    resultSheet.Cells[1, resBegin_FACTOR + k] = sdvhFactor.Values[k];

                //columnIndex++;
                columnIndex = resBegin_FACTOR + sdvhFactor.Values.Count;
                resultSheet.Cells[1, columnIndex] = COLUMN_STATUSROAD_NAME;
                int res_STATUSROAD = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_STATUSWEATHER_NAME;
                int res_STATUSWEATHER = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_LIGHT_NAME;
                int res_LIGHT = columnIndex;
                columnIndex++;
                resultSheet.Cells[1, columnIndex] = COLUMN_POINT_NAME;
                int res_POINT = columnIndex;
                //**//
                columnIndex++;
                int resBegin_BAD = columnIndex;
                for (int k = 0; k < sdvhBAD.Values.Count; k++)
                    resultSheet.Cells[1, resBegin_BAD + k] = sdvhBAD.Values[k];

                i = 2;
                int t;
                for (i = 2; sheet.Cells[i, 1].Value != null; i++)
                {
                    DateTime dt = DateTime.Parse(sheet.Cells[i, columnFromData].Text);
                    resultSheet.Cells[i, res_COLUMN_DAY] = dt.Day;
                    resultSheet.Cells[i, res_COLUMN_MONTH] = dt.Month;
                    resultSheet.Cells[i, res_COLUMN_WEEK] = dayOfWeekMassive.IndexOf(dt.ToString("dddd"));
                    resultSheet.Cells[i, res_COLUMN_PARTY] = PartyMassive.Contains(dt) ? 1 : 0;

                    DateTime tm = new DateTime();
                    var tmp = sheet.Cells[i, columnFromTime].Text.Split(':');
                    tm = tm.AddHours(Convert.ToInt32(tmp[0]));
                    tm = tm.AddMinutes(Convert.ToInt32(tmp[1]));
                    resultSheet.Cells[i, res_TIMEOFDAY_PARTY] = (double)tm.Hour + ((double)tm.Minute) / 60.0;//GetTimeOfDay(tm);

                    string typeDTP = sheet.Cells[i, columnFromTypeDTP].Text;
                    int typeDTPIndex = sdvhTypeDTP.Add(typeDTP);
                    resultSheet.Cells[i, res_TYPEDTP] = typeDTPIndex;

                    string road = sheet.Cells[i, columnFromRoad].Text;
                    int typeRoadIndex = sdvhRoad.Add(road);
                    resultSheet.Cells[i, res_ROAD] = typeRoadIndex;

                    resultSheet.Cells[i, res_KILOMETR] = sheet.Cells[i, columnFromKilometr].Text;

                    resultSheet.Cells[i, res_METR] = sheet.Cells[i, columnFromMetr].Text;

                    var newRowOfNDU = new bool[sdvhNDU.Values.Count];
                    var lstNDU = ParseString(sheet.Cells[i, columnFromNDU].Text);
                    foreach (var itm in lstNDU) newRowOfNDU[sdvhNDU[itm]] = true;
                    for (t = 0; t < newRowOfNDU.Length; t++) resultSheet.Cells[i, resBegin_NDU + t] = (newRowOfNDU[t] ? 1 : 0);

                    var newRowOfFactor = new bool[sdvhFactor.Values.Count];
                    var lstFactor = ParseString(sheet.Cells[i, columnFromFactors].Text);
                    foreach (var itm in lstFactor) newRowOfFactor[sdvhFactor[itm]] = true;
                    for (t = 0; t < newRowOfFactor.Length; t++) resultSheet.Cells[i, resBegin_FACTOR + t] = (newRowOfFactor[t] ? 1 : 0);

                    string stR = sheet.Cells[i, columnFromStatRoad].Text;
                    int stRINDX = sdvhStatRoad.Add(stR);
                    resultSheet.Cells[i, res_STATUSROAD] = stRINDX;

                    string stW = sheet.Cells[i, columnFromStatWeather].Text;
                    int stWINDX = sdvhStatWeather.Add(stW);
                    resultSheet.Cells[i, res_STATUSWEATHER] = stWINDX;

                    string light = sheet.Cells[i, columnFromLight].Text;
                    int lightINDX = sdvhLight.Add(light);
                    resultSheet.Cells[i, res_LIGHT] = lightINDX;

                    string point = sheet.Cells[i, columnFromPoint].Text;
                    resultSheet.Cells[i, res_POINT] = (point.Trim().ToLower() == "да" ? 1 : 0);

                    var newRowOfBAD = new bool[sdvhBAD.Values.Count];
                    var lstBAD = ParseString(sheet.Cells[i, columnFromBAD].Text);
                    foreach (var itm in lstBAD) newRowOfBAD[sdvhBAD[itm]] = true;
                    for (t = 0; t < newRowOfBAD.Length; t++) resultSheet.Cells[i, resBegin_BAD + t] = (newRowOfBAD[t] ? 1 : 0);

                    //if (i > 40) break;
                }

                ///////////----------------------------////////////////////

                int ii = 1;
                ExcplanationSheet.Cells[1, ii] = COLUMN_WEEK_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;//+ COLUMN_WEEK_NAME;
                for (int j = 0; j < dayOfWeekMassive.Count; j++)
                {
                    ExcplanationSheet.Cells[j + 2, ii] = dayOfWeekMassive[j];
                    ExcplanationSheet.Cells[j + 2, ii + 1] = j;
                }
                ii += 2;
                ExcplanationSheet.Cells[1, ii] = COLUMN_TYPEDTP_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;//+ COLUMN_TYPEDTP_NAME;
                ShowExplanation(sdvhTypeDTP, ii, ExcplanationSheet);
                ii += 2;
                ExcplanationSheet.Cells[1, ii] = COLUMN_ROAD_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;// + COLUMN_ROAD_NAME;
                ShowExplanation(sdvhRoad, ii, ExcplanationSheet);
                ii += 2;
                ExcplanationSheet.Cells[1, ii] = COLUMN_NDU_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;// + COLUMN_NDU_NAME;
                ShowExplanation(sdvhNDU, ii, ExcplanationSheet);
                ii += 2;
                ExcplanationSheet.Cells[1, ii] = COLUMN_FACTORS_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;// + COLUMN_FACTORS_NAME;
                ShowExplanation(sdvhFactor, ii, ExcplanationSheet);
                ii += 2;
                ExcplanationSheet.Cells[1, ii] = COLUMN_STATUSROAD_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;// + COLUMN_STATUSROAD_NAME;
                ShowExplanation(sdvhStatRoad, ii, ExcplanationSheet);
                ii += 2;
                ExcplanationSheet.Cells[1, ii] = COLUMN_STATUSWEATHER_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;// + COLUMN_STATUSWEATHER_NAME;
                ShowExplanation(sdvhStatWeather, ii, ExcplanationSheet);
                ii += 2;
                ExcplanationSheet.Cells[1, ii] = COLUMN_LIGHT_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;// + COLUMN_LIGHT_NAME;
                ShowExplanation(sdvhLight, ii, ExcplanationSheet);
                ii += 2;
                ExcplanationSheet.Cells[1, ii] = COLUMN_BAD_NAME;
                ExcplanationSheet.Cells[1, ii + 1] = SUFFIX_IN_EXPLANATION;// + COLUMN_BAD_NAME;
                ShowExplanation(sdvhBAD, ii, ExcplanationSheet);

                ExcplanationSheet.Columns.EntireColumn.AutoFit();
                resultSheet.Columns.EntireColumn.AutoFit();

                Console.WriteLine("Всего " + i);
            }
            catch (Exception e)
            {
                Console.WriteLine("!!!!");
                Console.WriteLine(e.Message);
            }


            ExcelApp.ScreenUpdating = true;
            ExcelApp.Interactive = true;
            ExcelApp.IgnoreRemoteRequests = false;
            ExcelApp.Visible = true;
            Console.WriteLine("Для выхода нажмиту любую клавишу...");
            Console.ReadKey();
        }
        public static string RmvExtrSpaces(string str)
        {
            if (str == null) return str;
            str = str.Trim();
            str = Regex.Replace(str, @"\s+", " ");
            return str;
        }
        public static List<string> ParseString(string str)
        {
            List<string> result = new List<string>();
            //while (!string.IsNullOrEmpty(str))
            //{
            //}
            //var arr = str.Split(',');
            result.AddRange(str.Split(',').Select(x => RmvExtrSpaces(x)).Where(x => !string.IsNullOrEmpty(x)));

            return result;
        }

        private static void ShowExplanation(StringDifferentValueHandler stringDifferentValueHandler, int indexBegin, Excel.Worksheet ExcplanationSheet)
        {
            for (int j = 0; j < stringDifferentValueHandler.Values.Count; j++)
            {
                ExcplanationSheet.Cells[j + 2, indexBegin] = stringDifferentValueHandler.Values[j];
                ExcplanationSheet.Cells[j + 2, indexBegin + 1] = j;
            }
        }

        class StringDifferentValueHandler
        {
            List<string> values;
            public int lastIndex;

            public List<string> Values
            {
                get { return values; }
            }

            public StringDifferentValueHandler()
            {
                values = new List<string>();
                lastIndex = -1;
            }

            public int Add(string newElem)
            {
                newElem = newElem.Trim().ToLower();
                lastIndex = values.IndexOf(newElem);
                if (lastIndex >= 0) return lastIndex;
                values.Add(newElem);
                lastIndex = values.Count - 1;
                return lastIndex;
            }

            public void Add(string[] newElems)
            {
                foreach (var item in newElems)
                {
                    this.Add(item);
                }
            }

            public int this[string key]
            {
                get { return values.IndexOf(values.FirstOrDefault(x => x.Trim().ToLower() == key.Trim().ToLower())); }
                //set { storage.SetAt(key, value); }
            }

        }

        //private static string GetTimeOfDay(DateTime dt)
        //{
        //    if (dt > DTMorning)
        //    {
        //        if (dt > DTDay)
        //        {
        //            if (dt > DTEvening)
        //            {
        //                return "Вечер";
        //            }
        //            else
        //            {
        //                return "День";
        //            }
        //        }
        //        else
        //        {
        //            return "Утро";
        //        }
        //    }
        //    else
        //    {
        //        return "Ночь";
        //    }
        //}
        //private static int GetTimeOfDay(DateTime dt)
        //{
        //    if (dt > DTMorning)
        //    {
        //        if (dt > DTDay)
        //        {
        //            if (dt > DTEvening)
        //            {
        //                return 3;
        //            }
        //            else
        //            {
        //                return 2;
        //            }
        //        }
        //        else
        //        {
        //            return 1;
        //        }
        //    }
        //    else
        //    {
        //        return 0;
        //    }
        //}
        //private static string GetTimeOfDay(int number)
        //{
        //    switch (number)
        //    {
        //        case 0: return "ночь";
        //        case 1: return "утро";
        //        case 2: return "день";
        //        case 3: return "вечер";
        //        default: return null;
        //    }
        //}

        private static int GetColumnIndex(Excel.Worksheet sheet, string columnName)
        {
            int clmnIndx = -1;
            for (int i = 1; sheet.Cells[1, i].Value != null; i++)
            {
                if (sheet.Cells[1, i].Text == columnName)
                {
                    clmnIndx = i;
                }
            }
            return clmnIndx;
        }
    }
}
