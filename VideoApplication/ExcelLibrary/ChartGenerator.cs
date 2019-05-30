using System;

using Microsoft.Office.Interop.Excel;

using static System.Runtime.InteropServices.Marshal;

namespace ExcelLibrary
{
	public class ChartGenerator : ExcelParent
	{
		private readonly Application Application;
		private readonly Workbook    Workbook;
		private          Worksheet   Worksheet;

		public ChartGenerator(string path)
		{
			Application = new Application();
			Application.Workbooks.Add(Template: path);
			Workbook  = Application.ActiveWorkbook;
			Worksheet = Workbook.Worksheets[Index: 1];

			GenerateDiagram(range: FindData());

			Workbook.SaveAs(Filename: @"D:\Chart.xlsx");
			Workbook.Close();
			FinalReleaseComObject(o: Worksheet);
			FinalReleaseComObject(o: Workbook);
			Application.Quit();
			FinalReleaseComObject(o: Application);
		}

		public void GenerateDiagram(Range range)
		{
			Workbook.Worksheets.Add();
			Worksheet = Workbook.Worksheets[Index: 2];

			ChartObjects xlCharts  = Worksheet.ChartObjects(Index: Type.Missing);
			ChartObject  myChart   = xlCharts.Add(Left: 10, Top: 80, Width: 300, Height: 250);
			Chart        chartPage = myChart.Chart;
			chartPage.ChartType = XlChartType.xl3DColumnClustered;
			chartPage.SetSourceData(Source: range);
		}

		public Range FindData() => Worksheet.UsedRange;
	}
}
