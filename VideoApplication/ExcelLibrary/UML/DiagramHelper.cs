using Microsoft.Office.Interop.Excel;

namespace ExcelLibrary.UML
{
	/// <summary>
	///   Клас для створення діаграм
	/// </summary>
	public class DiagramHelper
	{
		/// <summary>
		///   Конструктор для ініціалізації властивості
		/// </summary>
		/// <param name="range"></param>
		public DiagramHelper(Range range) => Range = range;

		public Range Range { get; set; }

		/// <summary>
		///   Метод для створення діаграми
		/// </summary>
		public void GenerateDiagram() { }
	}
}
