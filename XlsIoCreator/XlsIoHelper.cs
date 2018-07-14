using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsIoCreator
{
    public class XlsIoHelper
    {
        public static void SetValue(IWorksheet worksheet, int row, int col, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                worksheet.SetBlank(row, col);
            else
                worksheet.SetText(row, col, value);
        }

        public static void SetGeneral(IWorksheet worksheet, int row, int col, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                worksheet.SetBlank(row, col);
            else
                worksheet.SetText(row, col, value);

            worksheet[row, col].NumberFormat = "General";
        }

        public static void SetNumber(IWorksheet worksheet, int row, int col, double? value)
        {
            if (value.HasValue)
                worksheet.SetNumber(row, col, value.Value);
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetDecimal(IWorksheet worksheet, int row, int col, decimal? value)
        {
            if (value.HasValue)
                worksheet.SetNumber(row, col, (double)value.Value);
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetCurrency(IWorksheet worksheet, int row, int col, double? value)
        {
            if (value.HasValue)
            {
                worksheet.Range[row, col].NumberFormat = "$#,##0.00";
                worksheet.SetNumber(row, col, (double)value.Value / 1.1);
            }
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetCurrency(IWorksheet worksheet, int row, int col, decimal? value)
        {
            if (value.HasValue)
            {
                worksheet.Range[row, col].NumberFormat = "$#,##0.00";
                worksheet.SetNumber(row, col, (double)value.Value);
            }
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetPercentage(IWorksheet worksheet, int row, int col, decimal? value)
        {
            if (value.HasValue)
            {
                worksheet.Range[row, col].NumberFormat = "0.00%";
                worksheet.SetNumber(row, col, (double)value.Value);
            }
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetCurrencyGST(IWorksheet worksheet, int row, int col, decimal? value)
        {
            if (value.HasValue)
            {
                worksheet.Range[row, col].NumberFormat = "$#,##0.00";
                double numSet = (double)value.Value;
                numSet = numSet * 1.1;
                worksheet.SetNumber(row, col, numSet);
            }
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetDate(IWorksheet worksheet, int row, int col, DateTime? value)
        {
            if (value.HasValue)
            {
                worksheet.Range[row, col].NumberFormat = "d/m/yyyy";
                worksheet.Range[row, col].DateTime = value.Value.Date;
            }
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetDateTime(IWorksheet worksheet, int row, int col, DateTime? value)
        {
            if (value.HasValue)
            {
                worksheet.Range[row, col].NumberFormat = "d/m/yyyy h:mm:ss AM/PM";
                worksheet.Range[row, col].DateTime = value.Value;
            }
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetTime(IWorksheet worksheet, int row, int col, DateTime? value)
        {
            if (value.HasValue)
            {
                worksheet.Range[row, col].NumberFormat = "h:mm:ss AM/PM";
                worksheet.Range[row, col].DateTime = value.Value;
            }
            else
                worksheet.SetBlank(row, col);
        }

        public static void SetBoolean(IWorksheet worksheet, int row, int col, bool? value)
        {
            if (value.HasValue)
                worksheet.SetBoolean(row, col, value.Value);
            else
                worksheet.SetValue(row, col, "");
        }

        public static void SetHyperlink(IWorksheet worksheet, int row, int col, string value, string url)
        {
            var hyperlink = worksheet.HyperLinks.Add(worksheet[row, col]);
            hyperlink.Type = ExcelHyperLinkType.Url;
            hyperlink.Address = url;
            hyperlink.TextToDisplay = value;
        }
        public static void FormatWorksheet(IWorksheet worksheet, string title = "")
        {
            var col = 1;
            while (!string.IsNullOrWhiteSpace(worksheet[1, col].Value))
            {
                worksheet[1, col].CellStyle.Font.Bold = true;
                worksheet.AutofitColumn(col);
                col++;
            }

            var row = 1;
            while (!string.IsNullOrWhiteSpace(worksheet[row, 1].Value))
            {
                row++;
            }

            row++;
            row++;
            worksheet[row, 1].Value = "Information classification of this document is: RESTRICTED - Please refer R&S Security Manual for classification requirements.";
            worksheet[row, 1].CellStyle.Font.Size = 8;
            worksheet[row, 1].CellStyle.Font.Bold = true;
            worksheet[row, 1].CellStyle.Font.Color = ExcelKnownColors.Red;
            row++;
            worksheet[row, 1].Value = "Upon removal of this document from R&S information centre, it becomes uncontrolled.";
            worksheet[row, 1].CellStyle.Font.Size = 8;
            worksheet[row, 1].CellStyle.Font.Bold = true;
            worksheet[row, 1].CellStyle.Font.Color = ExcelKnownColors.Red;

            if (!string.IsNullOrWhiteSpace(title))
            {
                worksheet.InsertRow(1, 2);
                worksheet[1, 1].Value = title;
                worksheet[1, 1].CellStyle.Font.Size = 16;
                worksheet[1, 1].CellStyle.Font.Bold = true;
            }


        }

        public static void FormatWorksheet2(IWorksheet worksheet, string title = "")
        {
            var row = 2;
            while (!string.IsNullOrWhiteSpace(worksheet[row, 1].Value))
            {
                row++;
            }

            row++;
            row++;
            worksheet[row, 1].Value = "Information classification of this document is: RESTRICTED - Please refer R&S Security Manual for classification requirements.";
            worksheet[row, 1].CellStyle.Font.Size = 8;
            worksheet[row, 1].CellStyle.Font.Bold = true;
            worksheet[row, 1].CellStyle.Font.Color = ExcelKnownColors.Red;
            row++;
            worksheet[row, 1].Value = "Upon removal of this document from R&S information centre, it becomes uncontrolled.";
            worksheet[row, 1].CellStyle.Font.Size = 8;
            worksheet[row, 1].CellStyle.Font.Bold = true;
            worksheet[row, 1].CellStyle.Font.Color = ExcelKnownColors.Red;

            if (!string.IsNullOrWhiteSpace(title))
            {
                worksheet.InsertRow(1, 2);
                worksheet[1, 1].Value = title;
                worksheet[1, 1].CellStyle.Font.Size = 16;
                worksheet[1, 1].CellStyle.Font.Bold = true;
            }


        }

        public static void SetMonth(IWorksheet worksheet, int row, int col, DateTime? value)
        {
            if (value.HasValue)
            {
                worksheet.Range[row, col].NumberFormat = "mmm yyyy";
                worksheet.Range[row, col].DateTime = value.Value;
            }
            else
                worksheet.SetBlank(row, col);
        }
    }
}
