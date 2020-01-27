using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace XlsIoCreator
{
    public static class XlsIoCreator
    {
        private static IWorkbook GetWorkbook<T>(List<T> list, IApplication application)
        {

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Workbook.StandardFont = "Arial";
            worksheet.Workbook.StandardFontSize = 10;

            var row = 1;
            var col = 1;
            foreach (PropertyInfo prop in typeof(T).GetProperties())
            {
                var customAttributes = (DisplayAttribute[])prop.GetCustomAttributes(typeof(DisplayAttribute), true);
                if (customAttributes.Count() > 0)
                    worksheet[row, col++].Value = customAttributes[0].DisplayName;
                else
                    worksheet[row, col++].Value = prop.Name;
            }

            var numColumns = col - 1;
            worksheet.Rows[0].CellStyle.Font.Bold = true;
            row++;
            col = 1;

            foreach (var item in list)
            {
                col = 1;
                foreach (PropertyInfo prop in typeof(T).GetProperties())
                {
                    if (prop.PropertyType == typeof(string))
                        worksheet[row, col++].Value = Convert.ToString(prop.GetValue(item, null));
                    else if (prop.PropertyType == typeof(Nullable<int>) || prop.PropertyType == typeof(int))
                    {
                        var testInt = prop.GetValue(item, null);
                        if (testInt != null)
                            worksheet[row, col++].Number = Convert.ToInt32(testInt);
                        else
                            col++;
                    }
                    else if (prop.PropertyType == typeof(Nullable<DateTime>) || prop.PropertyType == typeof(DateTime))
                    {
                        var testDate = prop.GetValue(item, null);
                        if (testDate != null)
                        {
                            worksheet[row, col++].DateTime = Convert.ToDateTime(testDate);
                            var customAttributes = (DateFormatAttribute[])prop.GetCustomAttributes(typeof(DateFormatAttribute), true);
                            if (customAttributes.Count() > 0)
                            {
                                switch (customAttributes[0].DateFormat)
                                {
                                    case DateFormatAttribute.DateFormatEnum.DateOnly:
                                        worksheet[row, col - 1].NumberFormat = "d/mm/yyyy";
                                        break;
                                    case DateFormatAttribute.DateFormatEnum.TimeOnly:
                                        worksheet[row, col - 1].NumberFormat = "h:mm AM/PM";
                                        break;
                                    case DateFormatAttribute.DateFormatEnum.DateTime:
                                        worksheet[row, col - 1].NumberFormat = "d/mm/yyyy h:mm AM/PM";
                                        break;
                                }
                            }
                            else
                            {
                                if (Convert.ToDateTime(testDate).Ticks == 0)
                                    worksheet[row, col - 1].NumberFormat = "d/mm/yyyy";
                                else
                                    worksheet[row, col - 1].NumberFormat = "d/mm/yyyy h:mm AM/PM";
                            }
                        }
                        else
                            col++;
                    }
                    else if (prop.PropertyType == typeof(Nullable<decimal>) || prop.PropertyType == typeof(decimal))
                    {

                        var testDate = prop.GetValue(item, null);
                        if (testDate != null)
                        {
                            worksheet[row, col++].Number = Convert.ToDouble(testDate);
                            if (prop.GetCustomAttributes(typeof(CurrencyAttribute), true).Count() > 0)
                                worksheet[row, col - 1].NumberFormat = "$#,##0.00";

                        }
                        else
                            col++;
                    }
                }

                row++;
            }


            return workbook;
            //workbook.SaveAs(fileName);
        }

        public static void ToXlsIo<T>(this List<T> list, string fileName)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                var workbook = GetWorkbook(list, application);
                workbook.SaveAs(fileName);
            }
        }

        public static byte[] ToXlsIoBuffer<T>(this List<T> list)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                var workbook = GetWorkbook(list, application);
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.SaveAs(ms);
                    return ms.ToArray();
                }

            }
        }

    }

    public class CurrencyAttribute : Attribute
    {

    }

    public class DisplayAttribute : Attribute
    {
        public string DisplayName { get; set; }
    }

    public class DateFormatAttribute : Attribute
    {
        public enum DateFormatEnum
        {
            DateOnly,
            TimeOnly,
            DateTime
        }

        public DateFormatEnum DateFormat { get; set; }
    }
}
