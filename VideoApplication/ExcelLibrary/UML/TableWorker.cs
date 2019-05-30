using System.Windows.Forms;

namespace ExcelLibrary.UML
{
	/// <summary>
	///   Клас для роботи з таблицями в Excel
	/// </summary>
	public class TableWorker
	{
		/// <summary>
		///   Конструктор який ініціалізує властивість
		/// </summary>
		/// <param name="view"></param>
		public TableWorker(DataGridView view) => View = view;

		/// <summary>
		///   Властивість для здереження посилання на DataGridView
		/// </summary>
		public DataGridView View { get; set; }

		/// <summary>
		///   Метод для завантаження даних з листа Excel
		/// </summary>
		public void LoadFromExcel() { }

		/// <summary>
		///   Метод для збереження даних з DataGridView до Excel
		/// </summary>
		public void SaveToExcel() { }
	}
}
