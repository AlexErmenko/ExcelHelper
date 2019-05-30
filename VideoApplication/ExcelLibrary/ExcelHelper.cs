using System.IO;

using Microsoft.Office.Interop.Excel;

namespace ExcelLibrary
{
	public class ExcelHelper : ExcelParent
	{
		private WorksheetManager WorksheetManager;

		public Application Application { get; set; }
		public Workbook    Workbook    { get; set; }
		public Worksheet   Worksheet   { get; set; }
		public Sheets      Worksheets  { get; set; }

		public void Open(string path, bool visibleState = false)
		{
			if (!string.IsNullOrEmpty(value: path)
				&& File.Exists(path: path))
			{
				// Application.Workbooks.Open(Filename: path);
				// Application.Visible = visibleState;
				// Workbook            = Application.ActiveWorkbook;
			} else
				throw new FileNotFoundException(message: "Не вдалося відкрити заданий файл!");
		}

		public void Initialization()
		{
			Application = new Application();
			Application.Workbooks.Add();
			Workbook   = Application.ActiveWorkbook;
			Worksheets = Workbook.Worksheets;
			Worksheet  = Workbook.Worksheets[Index: 1];
		}

		public void SwitchWorksheet(int index)
		{
			Worksheet = Worksheets[Index: index];
		}

		public void AppendWorksheet()
		{
			Worksheets.Add(After: Worksheet);
		}
	}
}
