using Microsoft.Office.Interop.Excel;

namespace ExcelLibrary.UML
{
	/// <summary>
	///   Клас для роботи з листами
	/// </summary>
	public class WorksheetManager
	{
		/// <summary>
		///   Консструктор
		/// </summary>
		/// <param name="worksheet"></param>
		public WorksheetManager(Worksheet worksheet) => Worksheet = worksheet;

		/// <summary>
		///   Ініціалізація екземляра об'єкта листа
		/// </summary>
		public Worksheet Worksheet { get; set; }

		/// <summary>
		///   Метод для переходу до останього листа
		/// </summary>
		public void Last() { }

		/// <summary>
		///   Метод для переходу до першого листа
		/// </summary>
		public void First() { }

		/// <summary>
		///   Метод для переходу до попереднього листа
		/// </summary>
		public void Prev() { }

		/// <summary>
		///   Метод для переходу до наступного листа
		/// </summary>
		public void Next() { }
	}
}
