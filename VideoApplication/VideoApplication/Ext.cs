using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using static System.Runtime.InteropServices.Marshal;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelApplication
{
	public class Ext : IDisposable
	{
		public Ext(DataGridView view, string file)
		{
			View        = view;
			Application = new Application();
			Application.Workbooks.Add(Template: file);
			Workbook  = Application.ActiveWorkbook;
			Worksheet = Workbook.Worksheets[Index: 1];
			UsedRange = Worksheet.UsedRange;
		}

		public DataGridView View        { get; set; }
		public Application  Application { get; set; }
		public Workbook     Workbook    { get; set; }
		public Worksheet    Worksheet   { get; set; }
		public Range        UsedRange   { get; set; }

		public void Dispose()
		{
			ReleaseUnmanagedResources();
			GC.SuppressFinalize(obj: this);
		}

		private void ReleaseUnmanagedResources()
		{
			if (Application != null) FinalReleaseComObject(o: Application);
			if (Workbook    != null) FinalReleaseComObject(o: Workbook);
			if (Worksheet   != null) FinalReleaseComObject(o: Worksheet);
			if (UsedRange   != null) FinalReleaseComObject(o: UsedRange);
		}

		~Ext()
		{
			ReleaseUnmanagedResources();
		}

		public void LoadXLSX(string file)
		{
			int rowCount     = UsedRange.Rows.Count;
			int collumnCount = UsedRange.Columns.Count -1;

			var headerList = new List<string>();
			View.RowCount++;

			for (var i = 1; i <= collumnCount; i++)
			{
				try
				{
					dynamic collumnText =
							UsedRange?.Cells[RowIndex: 1, ColumnIndex: i]?.Value;

					if (collumnText == null) continue;

					View.ColumnCount++;
					View.Columns[i-1].HeaderText = collumnText;
				} catch (COMException e)
				{
					Console.WriteLine(value: e.Message);
				}
			}

			for (int i = 2; i <= rowCount; i++)
			{
				View.RowCount++;
				for (int j =1; j < collumnCount; j++)
				{
					dynamic cellValue = UsedRange?.Cells[i, j]?.Value;

					if (cellValue != null)
					{
						View[j - 1, i - 2].Value = cellValue;
					}
				}
			}
			// /*for (int i = View.RowCount; i < rowCount; i++)
			// {
			// 	View.RowCount++;
			//
			// 	for (var j = 1; j <= collumnCount; j++)
			// 	{
			// 		View.ColumnCount++;
			//
			// 		dynamic value =
			// 				UsedRange?.Cells[RowIndex: i + 1, ColumnIndex: j]?.Value;
			//
			// 		View[columnIndex: j - 1, rowIndex: i].Value = value;
			// 	}
			// }*/

			try
			{
				Workbook.Close();
				Application.Quit();
			} catch (COMException comException) { }
		}
	}
}
