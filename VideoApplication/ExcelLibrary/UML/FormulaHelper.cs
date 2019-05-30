using Microsoft.Office.Interop.Excel;

namespace ExcelLibrary.UML
{
	/// <summary>
	///   Клас для роботи з формулами
	/// </summary>
	public class FormulaHelper : ExcelParent
	{
		/// <summary>
		///   Конструктор для ініціалізації властивості
		/// </summary>
		/// <param name="range"></param>
		public FormulaHelper(Range range) => Range = range;

		/// <summary>
		///   Властивість для збереження посилання на діапазон з формулою
		/// </summary>
		public Range Range { get; set; }

		/// <summary>
		///   Перевірка корректності формули
		/// </summary>
		/// <param name="formula"></param>
		public void ValidateFormula(string formula) { }
	}
}
