using System;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
	internal static class Program
	{
		[STAThread]
		private static void Main()
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(defaultValue: false);
			Application.Run(new SqlQuery());
		}
	}
}
