using System;

using static System.Windows.Forms.Application;

namespace ExcelApplication
{
	internal static class Program
	{
		/// <summary>
		///   Главная точка входа для приложения.
		/// </summary>
		[STAThread]
		private static void Main()
		{
			EnableVisualStyles();
			SetCompatibleTextRenderingDefault(defaultValue: false);
			Run(mainForm: new Main());
		}
	}
}
