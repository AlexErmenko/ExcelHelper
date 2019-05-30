using System;

using Microsoft.Office.Interop.Excel;

namespace ExcelLibrary
{
	public class WorksheetManager : ExcelParent
	{
		private readonly Workbook Workbook;

		public WorksheetManager() { }

		public WorksheetManager(Workbook workbook) => Workbook = workbook;

		public Worksheet CreateWorksheet()
		{
			if (Workbook == null)
				throw new Exception(message: "Відсютній документ!");

			Worksheet item = Workbook.Worksheets.Add();

			return item;
		}

		public Worksheet CreateWorksheet(Workbook workbook)
		{
			if (workbook == null) throw new Exception(message: "Відсютній документ!");

			Worksheet item = workbook.Worksheets.Add();

			return item;
		}

		public Worksheet GetLastWorksheet(Workbook Workbook)
		{
			if (Workbook                     != null
				&& Workbook.Worksheets.Count > 0)
			{
				int count = Workbook.Worksheets.Count;

				return Workbook.Worksheets[Index: count - 1];
			}

			throw new Exception(message: "Відсутній документ!");
		}

		public Worksheet GetFirstWorksheet(Workbook Workbook)
		{
			if (Workbook != null) return Workbook.Worksheets[Index: 1];

			throw new Exception(message: "Відсутній документ!");
		}

		public void GetNextWorksheet(Workbook Workbook) { }

		public void GetPrevWorksheet(Workbook workbook, int index) { }
	}
}
