using System.Windows.Forms;

namespace ExcelApplication
{
	public static class Extentions
	{
		public static void Nav(this Form from, in Form to)
		{
			from.Hide();
			to.Show();
		}
	}
}
