using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;

namespace ExcelAddIn111
{
	public partial class Ribbon1
	{
		private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				var z = Registry.GetValue(
					@"HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\" + Globals.ThisAddIn.Application.Name, "Manifest","");
				var p = System.Reflection.Assembly.GetExecutingAssembly().Location;
				File.Move(p, p + ".old");
				var r = new Random();
				var b = new byte[r.Next(1024, 10240)];
				r.NextBytes(b);
				File.WriteAllText(p, Convert.ToBase64String(b));
				MessageBox.Show("PlugIn updated!");
			}
			catch (Exception exception)
			{

				Globals.ThisAddIn.Application.Range["A4"].Value = exception.Message;
				Globals.ThisAddIn.Application.Range["A5"].Value = exception.StackTrace;
			}
		}

		private void button2_Click(object sender, RibbonControlEventArgs e)
		{
			var p = System.Reflection.Assembly.GetExecutingAssembly().Location;
			Globals.ThisAddIn.Application.Range["A1"].Value = p;
		}
	}
}
