using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

using ExcelLibrary;

using Microsoft.Office.Interop.Excel;

using static System.Environment;
using static System.Environment.SpecialFolder;
using static System.IO.Directory;
using static System.String;
using static System.Windows.Forms.MessageBoxButtons;
using static System.Windows.Forms.MessageBoxIcon;

using static ExcelApplication.ModeWorksheet;
using static ExcelApplication.Properties.Resources;

using Application = Microsoft.Office.Interop.Excel.Application;
using Button = System.Windows.Forms.Button;
using Font = System.Drawing.Font;
using GroupBox = System.Windows.Forms.GroupBox;
using Point = System.Drawing.Point;

namespace ExcelApplication
{
	public partial class Main : Form
	{
		private static int i = 1;

		private readonly SaveFileDialog _dialog = new SaveFileDialog
												{
														DefaultExt = ".xlsx",
														Filter =
																"Excel Document 2007(*.xlsx)|*.xlsx",
														InitialDirectory =
																GetFolderPath(folder: MyDocuments)
												};

		private readonly ExcelHelper   Helper        = new ExcelHelper();
		private          ModeWorksheet ModeWorksheet = SelfCreate;

		public Main()
		{
			InitializeComponent();
		}

		private Range Range { get; set; }

		private DataGridViewColumn GenerateCollumn(string headerText) =>
				new DataGridViewColumn
				{
						HeaderText   = $"{headerText}",
						CellTemplate = new DataGridViewTextBoxCell()
				};

		/// <summary>
		///   Добавление колонки в таблицу
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void AppendColl_Click(object sender, EventArgs e)
		{
			TabPage      page = GetCurrentTab(tabControl: tabControl1);
			DataGridView data = GetDataGridViewByPage(page: page);

			string headerText =
					ColumnIndexToColumnLetter(colIndex: data.ColumnCount + 1);

			DataGridViewColumn column = GenerateCollumn(headerText: headerText);

			data.Columns.Add(dataGridViewColumn: column);
		}

		/// <summary>
		///   Метод для отримання поточного листа
		/// </summary>
		/// <param name="tabControl"></param>
		/// <returns>Возвращает текущую вкладку</returns>
		public TabPage GetCurrentTab(TabControl tabControl) => tabControl.SelectedTab;

		/// <summary>
		///   Метод для отримання літери по індексу
		/// </summary>
		/// <param name="colIndex">Индекс колонки</param>
		/// <returns>Символ из алфавита</returns>
		private static string ColumnIndexToColumnLetter(int colIndex)
		{
			int    div       = colIndex;
			string colLetter = Empty;
			var    mod       = 0;

			while (div > 0)
			{
				mod       = (div - 1) % 26;
				colLetter = (char) (65 + mod) + colLetter;
				div       = (div - mod) / 26;
			}

			return colLetter;
		}

		/// <summary>
		///   Сохранение в Excel таблицы
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void ToExcel_Click(object sender, EventArgs e)
		{
			TabPage      page = GetCurrentTab(tabControl: tabControl1);
			DataGridView view = GetDataGridViewByPage(page: page);

			if (view.RowCount       == 0
				|| view.ColumnCount == 0)
			{
				MessageBox.Show(text: "Відсутні дані для збереження!", caption: ProjectTitle,
								buttons: YesNo, icon: Warning);

				view.Select();

				return;
			}

			_dialog.FileOk += (o, args) => { };

			//TODO : поправить работу с диалоговомым окном
			string path = GetCurrentDirectory();

			TabPage      currentPage = tabControl1.SelectedTab;
			DataGridView data        = GetDataGridViewByPage(page: currentPage);

			if (data.Rows.Count != 0) data.SaveToExcel(path: GetCurrentDirectory());
			else MessageBox.Show(text: DataEmpty, caption: ProjectTitle);
		}

		/// <summary>
		///   Создание новой страницы по шаблону
		/// </summary>
		/// <param name="tabControl"></param>
		public void GenerateNewPage(TabControl tabControl)
		{
			#region RadioButtonLoad

			var radioButton = new RadioButton
							{
									AutoSize                = true,
									Location                = new Point(x: 6, y: 83),
									Name                    = "ModeLoadRadio",
									Size                    = new Size(width: 229, height: 21),
									TabIndex                = 1,
									TabStop                 = true,
									Text                    = "Завантаження зовнішніх даних",
									UseVisualStyleBackColor = true
							};

			radioButton.CheckedChanged += ModeLoadRadio_CheckedChanged;

			#endregion

			#region RadioButtonSelfCreate

			var radioButtonSelf = new RadioButton
								{
										AutoSize                = true,
										Location                = new Point(x: 6, y: 43),
										Name                    = "ModeWriteRadio",
										Size                    = new Size(width: 195, height: 21),
										TabIndex                = 0,
										TabStop                 = true,
										Text                    = "Власноручне заповнення",
										UseVisualStyleBackColor = true
								};

			radioButtonSelf.CheckedChanged += ModeWriteRadio_CheckedChanged;

			#endregion

			#region LoadButton

			var button = new Button
						{
								Location                = new Point(x: 677, y: 19),
								Margin                  = new Padding(all: 4),
								Name                    = "LoadTable",
								Size                    = new Size(width: 212, height: 35),
								TabIndex                = 5,
								Text                    = "Завантажити таблицю",
								UseVisualStyleBackColor = true
						};

			button.Click += LoadTable_Click;

			#endregion

			#region ToExcelButton

			var button1 = new Button
						{
								Location                = new Point(x: 356, y: 6),
								Margin                  = new Padding(all: 4),
								Name                    = "ToExcel",
								Size                    = new Size(width: 212, height: 35),
								TabIndex                = 4,
								Text                    = "Зберегти таблицю",
								UseVisualStyleBackColor = true
						};

			button1.Click += ToExcel_Click;

			#endregion

			#region AppendColumnButton

			var button2 = new Button
						{
								Location                = new Point(x: 357, y: 49),
								Margin                  = new Padding(all: 4),
								Name                    = "AppendColl",
								Size                    = new Size(width: 212, height: 35),
								TabIndex                = 3,
								Text                    = "Додати стовпець",
								UseVisualStyleBackColor = true
						};

			button2.Click += AppendColl_Click;

			#endregion

			#region GroupBox

			var groupBox = new GroupBox
							{
									Dock     = DockStyle.Left,
									Location = new Point(x: 3, y: 3),
									Name     = $"groupBox{i}",
									Size     = new Size(width: 243, height: 146),
									TabIndex = 6,
									TabStop  = false,
									Text     = "Режим роботи з листом"
							};

			groupBox.Controls.Add(value: radioButton);
			groupBox.Controls.Add(value: radioButtonSelf);

			#endregion

			#region DataGridView

			var data = new DataGridView
						{
								AllowUserToOrderColumns = true,
								ClipboardCopyMode = DataGridViewClipboardCopyMode
									.EnableAlwaysIncludeHeaderText,
								ColumnHeadersHeightSizeMode =
										DataGridViewColumnHeadersHeightSizeMode.AutoSize,
								ContextMenuStrip = contextMenuStrip1,
								Dock             = DockStyle.Bottom,
								Location         = new Point(x: 3, y: 149),
								Margin           = new Padding(all: 4),
								Name             = "SelfCreateTable",
								Size             = new Size(width: 891, height: 287),
								TabIndex         = 2
						};

			data.CellContentClick += SelfCreateTable_CellContentClick;

			#endregion

			#region TabPage

			TabPage tabPage = tabControl.SelectedTab;
			tabPage.Text = $"Лист {i}";

			tabPage.Font = new Font(family: FontFamily.GenericSansSerif, emSize: 10,
									style: FontStyle.Regular);

			var page = new TabPage
						{
								BorderStyle             = BorderStyle.FixedSingle,
								Location                = new Point(x: 4, y: 25),
								Name                    = $"tabPage{i}",
								Padding                 = new Padding(all: 3),
								TabIndex                = 0,
								Text                    = "+",
								Size                    = new Size(width: 899, height: 441),
								UseVisualStyleBackColor = true
						};

			tabPage.Controls.Add(value: groupBox);
			tabPage.Controls.Add(value: button);
			tabPage.Controls.Add(value: button1);
			tabPage.Controls.Add(value: button2);
			tabPage.Controls.Add(value: data);

			#endregion

			tabControl.TabPages.Add(value: page);
		}

		/// <summary>
		///   Метод для створення діагрм
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void Button1_Click(object sender, EventArgs e)
		{
			TabPage      page = GetCurrentTab(tabControl: tabControl1);
			DataGridView view = GetDataGridViewByPage(page: page);

			if (view.ColumnCount == 0
				|| view.RowCount == 0) { }

			// var dialog = new OpenFileDialog();
			// dialog.ShowDialog();
			// var generator = new ChartGenerator(path: dialog.FileName);
		}

		/// <summary>
		///   Добавление новой страницы
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void WorksheetSelectedIndexChanged(object sender, EventArgs e)
		{
			if (tabControl1.SelectedTab.Text == "+")
			{
				i++;
				GenerateNewPage(tabControl: tabControl1);
			}
		}

		/// <summary>
		///   Загрузка таблицы из Excel документа
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void LoadTable_Click(object sender, EventArgs e)
		{
			DialogResult result = OpenExcelDialog.ShowDialog();

			if (result == DialogResult.OK)
			{
				string       file = OpenExcelDialog.FileName;
				TabPage      page = GetCurrentTab(tabControl: tabControl1);
				DataGridView view = GetDataGridViewByPage(page: page);

				// view.LoadFromExcel(path: file);

				var ext = new Ext(view: view, file: file);
				ext.LoadXLSX(file: file);
				ext.Dispose();
			}
		}

		/// <summary>
		///   Обозначение диапазона для формулы в таблице
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void SelfCreateTable_CellContentClick(object sender, DataGridViewCellEventArgs e)
		{
			int rowIndex  = e.RowIndex;
			int collIndex = e.ColumnIndex;

			string path = GetCurrentDirectory();
			var    app  = new Application();
			app.Workbooks.Add(Template: path);
			Workbook  workbook  = app.ActiveWorkbook;
			Worksheet worksheet = workbook.Worksheets[Index: 1];

			Range = worksheet.Range[Cell1: rowIndex, Cell2: collIndex];

			Range.FormulaHidden = false;

			//range.Formula       = "SUM(A1:B1)";
			//range.FormulaHidden = false;
		}

		/// <summary>
		///   Выбор режима самостоятельного заполненния таблицы
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void ModeWriteRadio_CheckedChanged(object sender, EventArgs e)
		{
			TabPage      page = GetCurrentTab(tabControl: tabControl1);
			DataGridView view = GetDataGridViewByPage(page: page);
			view.Enabled = true;
			Button load = GetButtonLoadTable(page: page);
			load.Enabled = false;
			Button save = GetButtonSaveToExcel(page: page);
			save.Enabled = true;
			Button append = GetButtonAppendCollumn(page: page);
			append.Enabled = true;
		}

		/// <summary>
		///   Выбор режима загрузки внешних данных в таблицу
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void ModeLoadRadio_CheckedChanged(object sender, EventArgs e)
		{
			TabPage      page = GetCurrentTab(tabControl: tabControl1);
			DataGridView view = GetDataGridViewByPage(page: page);
			view.Enabled = true;
			Button load = GetButtonLoadTable(page: page);
			load.Enabled = true;
			Button save = GetButtonSaveToExcel(page: page);
			save.Enabled = true;
			Button append = GetButtonAppendCollumn(page: page);
			append.Enabled = false;
		}

		/// <summary>
		///   Метод для получения кнопки "сохранить" на текущей странице
		/// </summary>
		/// <param name="page">Текущая выбраная вкладка</param>
		/// <returns></returns>
		/// <exception cref="Exception"></exception>
		private Button GetButtonSaveToExcel(TabPage page)
		{
			foreach (object control in page.Controls)
			{
				if (control is Button btn
					&& btn.Name == "ToExcel")
					return btn;
			}

			throw new Exception();
		}

		/// <summary>
		///   Метод
		/// </summary>
		/// <param name="page">Текущая выбраная вкладка</param>
		/// <returns>Кнопка для добавления столбцов</returns>
		/// <exception cref="Exception"></exception>
		private Button GetButtonAppendCollumn(TabPage page)
		{
			foreach (object control in page.Controls)
			{
				if (control is Button btn
					&& btn.Name == "AppendColl")
					return btn;
			}

			throw new Exception();
		}

		/// <summary>
		///   Метод
		/// </summary>
		/// <param name="page">Текущая выбраная вкладка</param>
		/// <returns>Кнопка для загрузки данных из внешнего источника</returns>
		/// <exception cref="Exception"></exception>
		private Button GetButtonLoadTable(TabPage page)
		{
			foreach (object item in page.Controls)
			{
				if (item is Button btn
					&& btn.Name == "LoadTable")
					return btn;
			}

			throw new Exception(message: "Error!");
		}

		/// <summary>
		///   Метод для отримання таблиці з поточного листа
		/// </summary>
		/// <param name="page">Текущаяя страница</param>
		/// <returns>Таблицу на странице</returns>
		public DataGridView GetDataGridViewByPage(TabPage page)
		{
			foreach (Control control in page.Controls)
			{
				if (control is DataGridView data)
					return data;
			}

			return null;
		}

		private void Main_Load(object sender, EventArgs e)
		{
			Helper.Initialization();
			ModeWriteRadio.Checked = true;
		}

		private void OpenExcelDialog_FileOk(object sender, CancelEventArgs e)
		{
			// var page = GetCurrentTab(tabControl1);
			// var view = GetDataGridViewByPage(page);
			//
			// var file = OpenExcelDialog.FileName;
			//
			// view.LoadFromExcel(file);
		}
	}
}
