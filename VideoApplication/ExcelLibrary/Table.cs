using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using static System.Runtime.InteropServices.Marshal;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelLibrary
{
	// TODO: Использовать ф-и

	// TODO: Диаграмы?

	/// <summary>
	///   Класс для работы с таблицами Excel
	/// </summary>
	/// // TODO: Добавлять, редактировать, форматировать, удалять данные
	public static class Table
	{
		private static string filePath;

		/// <summary>
		///   Метод расширения для DataGridView
		/// </summary>
		/// <param name="data">Таблица для сохранения</param>
		public static void SaveToExcel(this DataGridView data, string path)
		{
			var application = new Application();
			application.Workbooks.Add();
			Workbook  workbook  = application.ActiveWorkbook;
			Worksheet worksheet = workbook.Worksheets[Index: 1];

			AddHeaderRow(worksheet: worksheet, view: data);
			AddCellsValue(worksheet: worksheet, data: data);

			SetAutoSize(worksheet: worksheet);
			workbook.SaveAs(Filename: path);
			workbook.Close();
			FinalReleaseComObject(o: worksheet);
			FinalReleaseComObject(o: workbook);
			application.Quit();
			FinalReleaseComObject(o: application);
		}

		private static void SetAutoSize(Worksheet worksheet)
		{
			for (var i = 1; i < worksheet.UsedRange.Columns.Count; i++)
				worksheet.Columns[RowIndex: i].AutoFit();
		}

		private static void AddCellsValue(Worksheet worksheet, DataGridView data)
		{
			for (var i = 0; i < data.RowCount; i++)
			{
				for (var j = 0; j < data.ColumnCount; j++)
				{
					worksheet.Cells[RowIndex: i + 1, ColumnIndex: j + 1] =
							$"{data[columnIndex: j, rowIndex: i].Value}";

					try
					{
						var result = (int) data[columnIndex: j, rowIndex: i].Value;
					} catch (Exception e)
					{
						Range range =
								worksheet.Cells[RowIndex: i + 1, ColumnIndex: j + 1];

						range.Formula = data[columnIndex: j, rowIndex: i].Value;
					}
				}
			}
		}

		private static void AddHeaderRow(Worksheet worksheet, DataGridView view)
		{
			for (var i = 0; i < view.ColumnCount; i++)
			{
				worksheet.Cells[RowIndex: 1, ColumnIndex: i + 1] =
						view.Columns[index: i]?.HeaderText;
			}
		}

		private static string ColumnIndexToColumnLetter(int colIndex)
		{
			int    div       = colIndex;
			string colLetter = string.Empty;
			var    mod       = 0;

			while (div > 0)
			{
				mod       = (div - 1) % 26;
				colLetter = (char) (65 + mod) + colLetter;
				div       = (div - mod) / 26;
			}

			return colLetter;
		}

		public static void LoadFromExcel(this DataGridView data, string path)
		{
			var application = new Application();
			application.Workbooks.Add(Template: path);
			Workbook  workbook     = application.ActiveWorkbook;
			Worksheet sheet        = workbook.Worksheets[Index: 1];
			int       rowsCount    = sheet.UsedRange.Rows.Count;
			int       collumnCount = sheet.UsedRange.Columns.Count;

			var   headerList = new List<string>();
			Range usedRange  = sheet.UsedRange;

			for (var i = 1; i <= collumnCount; i++)
			{
				try
				{
					dynamic collumnText =
							usedRange?.Cells[RowIndex: 1, ColumnIndex: i]?.Value;

					if (collumnText != null) headerList.Add(item: collumnText);
				} catch (COMException e)
				{
					Console.WriteLine(value: e.Message);
				}
			}

			/*for (int i = data.ColumnCount; i < collumnCount; i++)
				data.Columns.Add(columnName: $"{i}", headerText: headerList[index: i]);*/

			for (int i = data.RowCount; i < rowsCount; i++)
			{
				data.RowCount++;

				for (var j = 1; j <= collumnCount; j++)
				{
					data.ColumnCount++;

					dynamic value =
							usedRange?.Cells[RowIndex: i + 1, ColumnIndex: j]?.Value;

					data[columnIndex: j - 1, rowIndex: i].Value = value;
				}
			}

			workbook.Close();
			FinalReleaseComObject(o: sheet);
			FinalReleaseComObject(o: workbook);
			application.Quit();
			FinalReleaseComObject(o: application);
		}
	}

	// TODO: Создание листов и навигация 

	// TODO: Открывать Excel документы
	// TODO: Сохранять в формате xls , xlsx
}
