using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn111
{
	public partial class ThisAddIn
	{

		private void CheckUpdates()
		{
			//Thread.Sleep(10000);

			//Get the assembly information
		}


		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			Task.Factory.StartNew(CheckUpdates);
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
		}

		#region Код, автоматически созданный VSTO

		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += ThisAddIn_Startup;
			this.Shutdown += ThisAddIn_Shutdown;
		}

		#endregion
	}
}
