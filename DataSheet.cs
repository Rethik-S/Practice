using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Syncfusion.XlsIO;

namespace XLSIO_task
{
    public partial class Form1
    {
        private void PrepareDatasheet(IWorksheet dataSheet, IWorkbook workbook)
        {
            dataTable(dataSheet, workbook);
            DisplaySales(dataSheet);
            WrapText(dataSheet);
            ColorScaleFormat(dataSheet);
            DisplayHyperLink(dataSheet);
            NumberRangeFormat(dataSheet);
            DisplayList(dataSheet);
            DisplayShapes(dataSheet, workbook);
            RedTextConditionalFormat(dataSheet);
            ListConditionalFormat(dataSheet);


            dataSheet.SetColumnWidth(1, 8);
            dataSheet.SetRowHeight(21, 72);
            dataSheet.SetColumnWidth(4, 10);
            dataSheet.SetColumnWidth(2, 23);
            dataSheet.SetColumnWidth(3, 23);
            dataSheet.SetColumnWidth(5, 13.22);
            dataSheet.SetColumnWidth(8, 8.22);

        }
        private void dataTable(IWorksheet dataSheet, IWorkbook workbook)
        {
            //names
            IName name = workbook.Names.Add("Months");
            name.RefersToRange = dataSheet.Range["H1"];

            //data table
            dataSheet.Range["A1"].Formula = "=Months";
            dataSheet.Range["A2"].Text = "Jan";
            dataSheet.Range["A3"].Text = "Feb";
            dataSheet.Range["A4"].Text = "Mar";
            dataSheet.Range["A5"].Text = "Apr";
            dataSheet.Range["A6"].Text = "May";
            dataSheet.Range["A7"].Text = "June";
            dataSheet.Range["A8"].Text = "Jul";
            dataSheet.Range["A9"].Text = "Aug";
            dataSheet.Range["A10"].Text = "Sep";
            dataSheet.Range["A11"].Text = "Oct";
            dataSheet.Range["A12"].Text = "Nov";
            dataSheet.Range["A13"].Text = "Dec";

            dataSheet.Range["B1"].Text = "Internet Sales Amount";
            dataSheet.Range["B2"].Number = 226170;
            dataSheet.Range["B3"].Number = 212259;
            dataSheet.Range["B4"].Number = 181079;
            dataSheet.Range["B5"].Number = 188809;
            dataSheet.Range["B6"].Number = 198195;
            dataSheet.Range["B7"].Number = 235524;
            dataSheet.Range["B8"].Number = 185786;
            dataSheet.Range["B9"].Number = 196745;
            dataSheet.Range["B10"].Number = 164897;
            dataSheet.Range["B11"].Number = 175673;
            dataSheet.Range["B12"].Number = 212896;
            dataSheet.Range["B13"].Number = 325634;

            dataSheet.Range["C1"].Text = "Reseller Sales Amount";
            dataSheet.Range["C2"].Number = 170234;
            dataSheet.Range["C3"].Number = 189456;
            dataSheet.Range["C4"].Number = 168795;
            dataSheet.Range["C5"].Number = 143567;
            dataSheet.Range["C6"].Number = 163567;
            dataSheet.Range["C7"].Number = 163546;
            dataSheet.Range["C8"].Number = 143787;
            dataSheet.Range["C9"].Number = 149898;
            dataSheet.Range["C10"].Number = 153784;
            dataSheet.Range["C11"].Number = 164289;
            dataSheet.Range["C12"].Number = 172453;
            dataSheet.Range["C13"].Number = 223430;

            dataSheet.Range["D1"].Text = "Unit Price";
            dataSheet.Range["D2"].Number = 202;
            dataSheet.Range["D3"].Number = 204;
            dataSheet.Range["D4"].Number = 191;
            dataSheet.Range["D5"].Number = 223;
            dataSheet.Range["D6"].Number = 203;
            dataSheet.Range["D7"].Number = 185;
            dataSheet.Range["D8"].Number = 198;
            dataSheet.Range["D9"].Number = 196;
            dataSheet.Range["D10"].Number = 220;
            dataSheet.Range["D11"].Number = 218;
            dataSheet.Range["D12"].Number = 299;
            dataSheet.Range["D13"].Number = 185;

            dataSheet.Range["E1"].Text = "Customers";
            dataSheet.Range["E2"].Number = 1861;
            dataSheet.Range["E3"].Number = 1522;
            dataSheet.Range["E4"].Number = 1410;
            dataSheet.Range["E5"].Number = 1488;
            dataSheet.Range["E6"].Number = 1781;
            dataSheet.Range["E7"].Number = 2155;
            dataSheet.Range["E8"].Number = 1657;
            dataSheet.Range["E9"].Number = 1767;
            dataSheet.Range["E10"].Number = 1448;
            dataSheet.Range["E11"].Number = 1556;
            dataSheet.Range["E12"].Number = 1928;
            dataSheet.Range["E13"].Number = 2956;


            // dataSheet.Range["H1"] = "Months";
            dataSheet.Range["Months"].Text = "Months";
            dataSheet.Range["H2"].Text = "Jan";
            dataSheet.Range["H3"].Text = "Feb";
            dataSheet.Range["H4"].Text = "Mar";
            dataSheet.Range["H5"].Text = "Apr";
            dataSheet.Range["H6"].Text = "May";
            dataSheet.Range["H7"].Text = "June";
            dataSheet.Range["H8"].Text = "Jul";
            dataSheet.Range["H9"].Text = "Aug";
            dataSheet.Range["H10"].Text = "Sep";
            dataSheet.Range["H11"].Text = "Oct";
            dataSheet.Range["H12"].Text = "Nov";
            dataSheet.Range["H13"].Text = "Dec";


            //Total
            dataSheet.Range["A14"].Text = "Total";
            dataSheet.Range["B14"].Formula = "=SUM(B2:B13)";
            dataSheet.Range["C14"].Formula = "=SUM(C2:C13)";
            dataSheet.Range["D14"].Formula = "=AVERAGE(D2:D13)";
            dataSheet.Range["E14"].Formula = "=SUM(E2:E13)";

            //Sales titles
            dataSheet.Range["B17"].Text = "2018 Sales";
            dataSheet.Range["B18"].Text = "2018 Sales";
            dataSheet.Range["B19"].Text = "Gain %";

            dataSheet.Range["B2:D13"].NumberFormat = "_($* #,##0.00_)";

            IRange tableHeader = dataSheet.Range["A1:E1"];
            tableHeader.CellStyle.Color = Color.FromArgb(198, 224, 180);
            tableHeader.CellStyle.Font.Bold = true;

            IRange tableFooter = dataSheet.Range["A14:E14"];
            tableFooter.CellStyle.Color = Color.FromArgb(198, 224, 180);
            tableFooter.CellStyle.Font.Bold = true;

            IRange monthHeading = dataSheet.Range["H1"];
            monthHeading.CellStyle.Color = Color.FromArgb(198, 224, 180);
            monthHeading.CellStyle.Font.Bold = true;
            
            //border style
            IRange borderCells = dataSheet.Range["B2:B13"];
            borderCells.BorderAround(ExcelLineStyle.Thin);
            borderCells.BorderInside(ExcelLineStyle.Thin);
        }
        private void DisplaySales(IWorksheet dataSheet)
        {
            //Sales values
            IRange salesValue = dataSheet.Range["C17"];
            salesValue.Number = 3845634;
            salesValue.CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;

            dataSheet.Range["C18"].Formula = "=B14+C14";
            dataSheet.Range["C17:C18"].NumberFormat = "_($* #,##0.00_)";

            IRange gain = dataSheet.Range["C19"];
            gain.Formula = "=(C18-C17)/10000000";
            gain.NumberFormat = "#.##%";
        }
        private void WrapText(IWorksheet dataSheet)
        {
            IRange wrapCell = dataSheet.Range["A21"];
            wrapCell.Text = "Syncfusion\nXlsIO\nDocIO\nPDF";
            wrapCell.WrapText = true;
            // wrapCell.RowHeight = 67;
            wrapCell.ColumnWidth = 9;
        }
        private void ColorScaleFormat(IWorksheet dataSheet)
        {
            //Create color scale for the data in specified range
            IConditionalFormats conditionalFormats = dataSheet.Range["D2:D13"].ConditionalFormats;
            IConditionalFormat conditionalFormat = conditionalFormats.AddCondition();
            conditionalFormat.FormatType = ExcelCFType.ColorScale;
            IColorScale colorScale = conditionalFormat.ColorScale;

            //Sets 2 - color scale and its constraints
            colorScale.SetConditionCount(2);
            colorScale.Criteria[0].FormatColorRGB = Color.FromArgb(255, 113, 40);
            colorScale.Criteria[0].Type = ConditionValueType.LowestValue;

            colorScale.Criteria[1].FormatColorRGB = Color.FromArgb(255, 239, 156);
            colorScale.Criteria[1].Type = ConditionValueType.HighestValue;
        }
        private void DisplayHyperLink(IWorksheet dataSheet)
        {
            //Create a hyperlink for a website
            IHyperLink hyperlink = dataSheet.HyperLinks.Add(dataSheet.Range["E17"]);
            hyperlink.Type = ExcelHyperLinkType.Url;
            hyperlink.Address = "https://help.syncfusion.com/file-formats/xlsio/overview";
            hyperlink.TextToDisplay = "SyncfusionXlsIO";
        }
        private void NumberRangeFormat(IWorksheet dataSheet)
        {
            //Applying conditional formatting 
            IConditionalFormats condition = dataSheet.Range["E2:E13"].ConditionalFormats;
            IConditionalFormat condition1 = condition.AddCondition();

            //Represents conditional format rule that the value in target range should be between 1500 and 1750
            condition1.FormatType = ExcelCFType.CellValue;
            condition1.Operator = ExcelComparisonOperator.Between;
            condition1.FirstFormula = "1500";
            condition1.SecondFormula = "1750";

            //Setting back color and font style to be applied for target range
            condition1.BackColorRGB = Color.FromArgb(0, 176, 80);
            condition1.FontColor = ExcelKnownColors.Red;
            condition1.IsBold = true;
        }
        private void DisplayList(IWorksheet dataSheet)
        {
            IDataValidation validation = dataSheet.Range["H2:H13"].DataValidation;
            validation.ListOfValues = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "June", "Jul", "Aug", "Sep", "Nov", "Dec" };
        }
        private void DisplayShapes(IWorksheet dataSheet, IWorkbook workbook)
        {

            ITextBoxShape shape1 = dataSheet.TextBoxes.AddTextBox(22, 1, 30, 130);

            //arrow
            IShape ARROW = dataSheet.Shapes.AddAutoShapes(AutoShapeType.Line, 23, 1, 73, 0);

            //arrow styles
            IShapeLineFormat lineFormat = ARROW.Line;
            lineFormat.BeginArrowheadWidth = ExcelShapeArrowWidth.ArrowHeadNarrow;
            lineFormat.EndArrowHeadStyle = ExcelShapeArrowStyle.LineArrow;
            lineFormat.ForeColor = Color.FromArgb(68, 114, 196);

            ARROW.Left = 60;
            ARROW.Top = shape1.Top + shape1.Height + 3;
            ARROW.Line.Weight = 0.1;

            ITextBoxShape shape2 = dataSheet.TextBoxes.AddTextBox(27, 1, 50, 120);
            shape2.Text = "XlsIO";

            //Set rich text
            IRichTextString richText = shape1.RichText;
            richText.Text = "Syncfusion";

            //Set font
            IFont font = workbook.CreateFont();
            font.Color = ExcelKnownColors.White;
            shape1.Fill.ForeColor = Color.FromArgb(68, 114, 196);
            richText.SetFont(0, 10, font);
            shape1.Line.DashStyle = ExcelShapeDashLineStyle.Solid;
            shape1.Line.ForeColor = Color.Blue;
            shape1.Line.Weight = 0.5;
        }
        private void RedTextConditionalFormat(IWorksheet dataSheet)
        {
            //months conditional formatting
            IConditionalFormats formats = dataSheet["A1:A13"].ConditionalFormats;

            IConditionalFormat format1 = formats.AddCondition();

            format1.FormatType = ExcelCFType.SpecificText;
            format1.FontColor = ExcelKnownColors.Red;

            format1.Text = "a";
        }
        private void ListConditionalFormat(IWorksheet dataSheet)
        {
            IConditionalFormats monthsFormats = dataSheet["H1:H13"].ConditionalFormats;
            //H column month
            IConditionalFormat format2 = monthsFormats.AddCondition();

            format2.FormatType = ExcelCFType.SpecificText;

            format2.BackColorRGB = Color.FromArgb(0, 176, 80);

            format2.BottomBorderStyle = ExcelLineStyle.Thin;
            format2.TopBorderStyle = ExcelLineStyle.Thin;
            format2.LeftBorderStyle = ExcelLineStyle.Thin;
            format2.RightBorderStyle = ExcelLineStyle.Thin;

            format2.BottomBorderColor = ExcelKnownColors.Red;
            format2.TopBorderColor = ExcelKnownColors.Red;
            format2.LeftBorderColor = ExcelKnownColors.Red;
            format2.RightBorderColor = ExcelKnownColors.Red;

            format2.Text = "j";
        }

    }
}