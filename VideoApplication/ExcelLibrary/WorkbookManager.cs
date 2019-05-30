using System;
using System.IO;

using Microsoft.Office.Interop.Excel;

namespace ExcelLibrary
{
	public class WorkbookManager : ExcelParent
	{
		public WorkbookManager() { }

		public WorkbookManager(Application application) => Application = application;
		public Application Application { get; }

		public Workbook Workbook { get; set; }

		public void CreateWorkbook()
		{
			Workbook = Application?.Workbooks.Add();
		}

		public void CreateWorkbook(Application application)
		{
			if (application == null) throw new NullReferenceException(message: nameof(application));
		}

		public void SaveAsXlsx(string path)
		{
			string file     = Path.GetFileNameWithoutExtension(path: path);
			string fullPath = file + ".xlsx";
			Workbook?.SaveAs(Filename: fullPath);
		}

		public void SaveAsXlsx(Workbook workbook, string path)
		{
			string file     = Path.GetFileNameWithoutExtension(path: path);
			string fullPath = file + ".xlsx";
			workbook?.SaveAs(Filename: fullPath);
		}

		public void SaveAsXls(string path)
		{
			string file     = Path.GetFileNameWithoutExtension(path: path);
			string fullPath = file + ".xls";
			Workbook?.SaveAs(Filename: fullPath);
		}

		public void SaveAsXls(Workbook workbook, string path)
		{
			string file     = Path.GetFileNameWithoutExtension(path: path);
			string fullPath = file + ".xls";
			workbook?.SaveAs(Filename: fullPath);
		}
	}
}
