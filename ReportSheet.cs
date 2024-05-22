using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Syncfusion.XlsIO;

namespace XLSIO_task
{
    public partial class Form1
    {
        private void PrepareReportSheet(IWorksheet dataSheet, IWorksheet salesReportSheet, IWorkbook workbook)
        {
            //sale title
            IRange title = salesReportSheet.Range["A1"];
            title.Formula = "=K1";
            title.CellStyle.Font.Size = 14;
            title.CellStyle.Font.Bold = true;
            title.CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            title.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            title.CellStyle.Color = Color.FromArgb(155, 194, 230);
            salesReportSheet.Range["A1:G1"].Merge();
            
            BarChart(salesReportSheet, dataSheet);
            ScatterLineChart(salesReportSheet, dataSheet);
            SalesTable(salesReportSheet, workbook);
            DisplayBlocks(salesReportSheet, workbook);

            //width
            salesReportSheet.SetRowHeight(1, 22.5);
            salesReportSheet.SetRowHeight(16, 3);
            salesReportSheet.SetColumnWidth(1, 3);
            salesReportSheet.SetColumnWidth(2, 11);
            salesReportSheet.SetColumnWidth(3, 11);
            salesReportSheet.SetColumnWidth(4, 0.81);
            salesReportSheet.SetColumnWidth(5, 11);
            salesReportSheet.SetColumnWidth(6, 11);
            salesReportSheet.SetColumnWidth(7, 3);

            salesReportSheet.SetRowHeight(14, 21);
            salesReportSheet.SetRowHeight(17, 21);

            salesReportSheet.SetRowHeight(15, 15);
            salesReportSheet.SetRowHeight(18, 15);

            salesReportSheet.SetColumnWidth(11, 11.78);
            salesReportSheet.SetColumnWidth(12, 9);
            salesReportSheet.SetColumnWidth(13, 9);
            

            salesReportSheet.IsGridLinesVisible = false;

        }
        private void BarChart(IWorksheet salesReportSheet, IWorksheet dataSheet)
        {
            //Create a Chart
            IChartShape chart1 = salesReportSheet.Charts.Add();
            chart1.Name = "Chart1";
            //Set Chart Type
            chart1.ChartType = ExcelChartType.Bar_Clustered;

            chart1.TopRow = 2;
            chart1.LeftColumn = 1;
            chart1.RightColumn = 8;
            chart1.BottomRow = 13;

            IChartSerie product1 = chart1.Series.Add("Internet Sales Amount");
            product1.Values = dataSheet.Range["B2:B13"];
            product1.CategoryLabels = dataSheet.Range["A2:A13"];

            chart1.PrimaryValueAxis.NumberFormat = "$#,###";
            chart1.PrimaryValueAxis.HasMajorGridLines = false;

            chart1.ChartTitle = "Internet Sales Amount";
            chart1.ChartTitleArea.Size = 18;
            chart1.PrimaryValueAxis.Title = "Axis Title";
            chart1.PrimaryCategoryAxis.Title = "Axis Title";
            chart1.PrimaryCategoryAxis.TextRotationAngle = -45;
            chart1.Legend.TextArea.Size = 12;


            //Axis title area text angle rotation
            chart1.PrimaryCategoryAxis.TitleArea.TextRotationAngle = -90;
            chart1.ChartArea.Border.LinePattern = ExcelChartLinePattern.None;
            chart1.Legend.TextArea.Size = 10;

            chart1.Legend.Position = ExcelLegendPosition.Bottom;

            // Customize the appearance of the bars
            foreach (IChartSerie series in chart1.Series)
            {
                series.SerieFormat.Fill.ForeColor = Color.FromArgb(68, 114, 196); // Set bar color to blue
                series.SerieFormat.Fill.Transparency = 0; // Set transparency to fully opaque
            }

        }
        private void ScatterLineChart(IWorksheet salesReportSheet, IWorksheet dataSheet)
        {
            //Create a Chart
            IChartShape chart2 = salesReportSheet.Charts.Add();
            chart2.Name = "Chart2";
            chart2.PrimaryValueAxis.HasMajorGridLines = false;

            //Set Chart Type
            chart2.ChartType = ExcelChartType.Scatter_Line_Markers;
            chart2.Legend.TextArea.Size = 9;
            chart2.Legend.Position = ExcelLegendPosition.Bottom;

            chart2.PrimaryValueAxis.HasMajorGridLines = false;

            chart2.PrimaryValueAxis.NumberFormat = "$#,###";

            chart2.TopRow = 20;
            chart2.LeftColumn = 1;
            chart2.RightColumn = 8;
            chart2.BottomRow = 32;

            //Set Chart Title
            chart2.ChartTitle = "Internet Sales vs Reseller Sales";

            //Set first serie
            IChartSerie productA = chart2.Series.Add("Internet Sales Amount");
            productA.Values = dataSheet.Range["B2:B13"];
            productA.CategoryLabels = dataSheet.Range["A2:A13"];

            //Set second serie
            IChartSerie productB = chart2.Series.Add("Reseller Sales Amount");
            productB.Values = dataSheet.Range["C2:C13"];
            productB.CategoryLabels = dataSheet.Range["A2:A13"];

            chart2.ChartArea.Border.LinePattern = ExcelChartLinePattern.None;
        }
        private void SalesTable(IWorksheet salesReportSheet, IWorkbook workbook)
        {
            //table
            salesReportSheet.Range["K1"].Text = "Yearly Sales";
            salesReportSheet.Range["K2"].Number = 400;
            salesReportSheet.Range["K3"].Number = 300;
            salesReportSheet.Range["K4"].Number = 500;
            salesReportSheet.Range["K5"].Number = 50;
            salesReportSheet.Range["K6"].Number = 200;



            salesReportSheet.Range["L1"].Text = "Expense";
            salesReportSheet.Range["L2"].Number = 36;
            salesReportSheet.Range["L3"].Number = 75;
            salesReportSheet.Range["L4"].Number = 75;
            salesReportSheet.Range["L5"].Number = 24;
            salesReportSheet.Range["L6"].Number = 64;

            salesReportSheet.Range["M1"].Text = "Growth";
            salesReportSheet.Range["M2"].Number = 45;
            salesReportSheet.Range["M3"].Number = 32;
            salesReportSheet.Range["M4"].Number = 64;
            salesReportSheet.Range["M5"].Number = 52;
            salesReportSheet.Range["M6"].Number = 75;
            salesReportSheet.Range["M2:M6"].NumberFormat = ".00%";


            IListObject table = salesReportSheet.ListObjects.Create("Table1", salesReportSheet["K1:M6"]);
            table.ShowTotals = true;
            table.Columns[0].TotalsRowLabel = "Total";
            table.Columns[1].TotalsCalculation = ExcelTotalsCalculation.None;
            table.Columns[2].TotalsCalculation = ExcelTotalsCalculation.Sum;


            //Apply custom table style
            ITableStyles tableStyles = workbook.TableStyles;
            ITableStyle tableStyle = tableStyles.Add("Table Style 1");
            ITableStyleElements tableStyleElements = tableStyle.TableStyleElements;


            ITableStyleElement tableStyleElement1 = tableStyleElements.Add(ExcelTableStyleElementType.FirstColumn);
            tableStyleElement1.BackColorRGB = Color.FromArgb(150, 198, 206);

            ITableStyleElement tableStyleElement2 = tableStyleElements.Add(ExcelTableStyleElementType.HeaderRow);
            tableStyleElement2.BackColorRGB = Color.FromArgb(0, 176, 240);

            ITableStyleElement tableStyleElement3 = tableStyleElements.Add(ExcelTableStyleElementType.TotalRow);
            tableStyleElement3.BackColorRGB = Color.FromArgb(0, 112, 192);

            ITableStyleElement tableStyleElement5 = tableStyleElements.Add(ExcelTableStyleElementType.WholeTable);
            tableStyleElement5.BackColorRGB = Color.FromArgb(0, 176, 80);
            tableStyleElement5.FontColorRGB = Color.FromArgb(0, 112, 192);

            ITableStyleElement tableStyleElement4 = tableStyleElements.Add(ExcelTableStyleElementType.LastColumn);
            tableStyleElement4.BackColorRGB = Color.FromArgb(208, 196, 154);

            table.TableStyleName = tableStyle.Name;
            table.ShowFirstColumn = true;
            table.ShowLastColumn = true;
        }
        private void DisplayBlocks(IWorksheet salesReportSheet, IWorkbook workbook)
        {
            IStyle style1 = workbook.Styles.Add("style1");
            style1.BeginUpdate();
            style1.Color = Color.FromArgb(155, 194, 230);
            style1.Font.Bold = true;
            style1.Font.Size = 18;
            style1.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            style1.VerticalAlignment = ExcelVAlign.VAlignCenter;
            style1.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
            style1.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
            style1.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
            style1.Borders[ExcelBordersIndex.EdgeTop].ColorRGB = Color.FromArgb(192, 192, 192);
            style1.Borders[ExcelBordersIndex.EdgeLeft].ColorRGB = Color.FromArgb(192, 192, 192);
            style1.Borders[ExcelBordersIndex.EdgeRight].ColorRGB = Color.FromArgb(192, 192, 192);
            style1.EndUpdate();

            IStyle style2 = workbook.Styles.Add("style2");
            style2.BeginUpdate();
            style2.Color = Color.FromArgb(244, 176, 132);
            style2.Font.Bold = true;
            style2.Font.Size = 18;
            style2.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            style2.VerticalAlignment = ExcelVAlign.VAlignCenter;
            style2.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
            style2.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
            style2.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
            style2.Borders[ExcelBordersIndex.EdgeTop].ColorRGB = Color.FromArgb(192, 192, 192);
            style2.Borders[ExcelBordersIndex.EdgeLeft].ColorRGB = Color.FromArgb(192, 192, 192);
            style2.Borders[ExcelBordersIndex.EdgeRight].ColorRGB = Color.FromArgb(192, 192, 192);
            style2.EndUpdate();

            IStyle style3 = workbook.Styles.Add("style3");
            style3.BeginUpdate();
            style3.Color = Color.FromArgb(255, 217, 102);
            style3.Font.Bold = true;
            style3.Font.Size = 18;
            style3.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            style3.VerticalAlignment = ExcelVAlign.VAlignCenter;
            style3.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
            style3.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
            style3.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
            style3.Borders[ExcelBordersIndex.EdgeTop].ColorRGB = Color.FromArgb(192, 192, 192);
            style3.Borders[ExcelBordersIndex.EdgeLeft].ColorRGB = Color.FromArgb(192, 192, 192);
            style3.Borders[ExcelBordersIndex.EdgeRight].ColorRGB = Color.FromArgb(192, 192, 192);
            style3.EndUpdate();

            IStyle style4 = workbook.Styles.Add("style4");
            style4.BeginUpdate();
            style4.Color = Color.FromArgb(169, 208, 142);
            style4.Font.Bold = true;
            style4.Font.Size = 18;
            style4.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            style4.VerticalAlignment = ExcelVAlign.VAlignCenter;
            style4.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
            style4.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
            style4.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
            style4.Borders[ExcelBordersIndex.EdgeTop].ColorRGB = Color.FromArgb(192, 192, 192);
            style4.Borders[ExcelBordersIndex.EdgeLeft].ColorRGB = Color.FromArgb(192, 192, 192);
            style4.Borders[ExcelBordersIndex.EdgeRight].ColorRGB = Color.FromArgb(192, 192, 192);
            style4.EndUpdate();

            IStyle style5 = workbook.Styles.Add("style5");
            style5.BeginUpdate();
            style5.Color = Color.FromArgb(155, 194, 230);
            style5.Font.Size = 11;
            style5.Font.Color = ExcelKnownColors.Grey_40_percent;
            style5.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            style5.VerticalAlignment = ExcelVAlign.VAlignCenter;
            style5.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
            style5.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
            style5.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
            style5.Borders[ExcelBordersIndex.EdgeBottom].ColorRGB = Color.FromArgb(192, 192, 192);
            style5.Borders[ExcelBordersIndex.EdgeLeft].ColorRGB = Color.FromArgb(192, 192, 192);
            style5.Borders[ExcelBordersIndex.EdgeRight].ColorRGB = Color.FromArgb(192, 192, 192);
            style5.EndUpdate();

            IStyle style6 = workbook.Styles.Add("style6");
            style6.BeginUpdate();
            style6.Color = Color.FromArgb(244, 176, 132);
            style6.Font.Color = ExcelKnownColors.Grey_40_percent;
            style6.Font.Size = 11;
            style6.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            style6.VerticalAlignment = ExcelVAlign.VAlignCenter;
            style6.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
            style6.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
            style6.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
            style6.Borders[ExcelBordersIndex.EdgeBottom].ColorRGB = Color.FromArgb(192, 192, 192);
            style6.Borders[ExcelBordersIndex.EdgeLeft].ColorRGB = Color.FromArgb(192, 192, 192);
            style6.Borders[ExcelBordersIndex.EdgeRight].ColorRGB = Color.FromArgb(192, 192, 192);
            style6.EndUpdate();

            IStyle style7 = workbook.Styles.Add("style7");
            style7.BeginUpdate();
            style7.Color = Color.FromArgb(255, 217, 102);
            style7.Font.Size = 11;
            style7.Font.Color = ExcelKnownColors.Grey_40_percent;
            style7.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            style7.VerticalAlignment = ExcelVAlign.VAlignCenter;
            style7.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
            style7.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
            style7.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
            style7.Borders[ExcelBordersIndex.EdgeBottom].ColorRGB = Color.FromArgb(192, 192, 192);
            style7.Borders[ExcelBordersIndex.EdgeLeft].ColorRGB = Color.FromArgb(192, 192, 192);
            style7.Borders[ExcelBordersIndex.EdgeRight].ColorRGB = Color.FromArgb(192, 192, 192);
            style7.EndUpdate();

            IStyle style8 = workbook.Styles.Add("style8");
            style8.BeginUpdate();
            style8.Color = Color.FromArgb(169, 208, 142);
            style8.Font.Color = ExcelKnownColors.Grey_40_percent;
            style8.Font.Size = 11;
            style8.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            style8.VerticalAlignment = ExcelVAlign.VAlignCenter;
            style8.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
            style8.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
            style8.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
            style8.Borders[ExcelBordersIndex.EdgeBottom].ColorRGB = Color.FromArgb(192, 192, 192);
            style8.Borders[ExcelBordersIndex.EdgeLeft].ColorRGB = Color.FromArgb(192, 192, 192);
            style8.Borders[ExcelBordersIndex.EdgeRight].ColorRGB = Color.FromArgb(192, 192, 192);
            style8.EndUpdate();

            IRange salesAmountValue = salesReportSheet.Range["B14"];
            salesAmountValue.Formula = "=Data!E14";
            salesAmountValue.CellStyle = style1;
            salesAmountValue.NumberFormat = "$#,###.00";
            salesAmountValue.CellStyle.Font.Color = ExcelKnownColors.Red;

            IRange salesAmountTitle = salesReportSheet.Range["B15"];
            salesAmountTitle.Text = "Sales Amount";
            salesAmountTitle.CellStyle = style5;

            salesReportSheet.Range["B15:C15"].Merge();
            salesReportSheet.Range["B14:C14"].Merge();


            IRange AVGUnitPriceValue = salesReportSheet.Range["E14"];
            AVGUnitPriceValue.Formula = "=Data!D14";
            AVGUnitPriceValue.CellStyle = style2;
            AVGUnitPriceValue.NumberFormat = "$#,###.00";
            AVGUnitPriceValue.CellStyle.Font.Underline = ExcelUnderline.Single;
            AVGUnitPriceValue.CellStyle.Font.Italic = true;

            IRange AVGUnitPriceTitle = salesReportSheet.Range["E15"];
            AVGUnitPriceTitle.Text = "Average Unit Price";
            AVGUnitPriceTitle.CellStyle = style6;

            salesReportSheet.Range["E14:F14"].Merge();
            salesReportSheet.Range["E15:F15"].Merge();

            IRange GrossTitle = salesReportSheet.Range["B18"];
            IRange GrossValue = salesReportSheet.Range["B17"];
            GrossValue.Formula = "=Data!C19";
            GrossValue.CellStyle = style3;
            GrossValue.NumberFormat = "#.##%";

            GrossTitle.Text = "Gross Profit Margin";
            GrossTitle.CellStyle = style7;

            salesReportSheet.Range["B17:C17"].Merge();
            salesReportSheet.Range["B18:C18"].Merge();

            IRange CustomerValue = salesReportSheet.Range["E17"];
            IRange CustomerCount = salesReportSheet.Range["E18"];
            CustomerValue.Formula = "=Data!E14";
            CustomerValue.CellStyle = style4;
            CustomerValue.NumberFormat = "#,###";
            CustomerValue.CellStyle.Font.RGBColor = Color.FromArgb(0, 176, 80);

            CustomerCount.Text = "Customer Count";
            CustomerCount.CellStyle = style8;

            salesReportSheet.Range["E17:F17"].Merge();
            salesReportSheet.Range["E18:F18"].Merge();
        }

    }
}