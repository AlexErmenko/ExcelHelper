using System;
using System.Windows.Forms;

namespace ExcelApplication
{
	public partial class Start : Form
	{
		public Start()
		{
			InitializeComponent();
		}

		public void GetStart(object _, EventArgs args) => this.Nav(to: new Main());
	}
}
