using DiscordRPC;

using Microsoft.Office.Interop.Excel;

using System;

namespace DiscordForExcel
{
	public partial class ThisAddIn
	{
		public DiscordRpcClient Client;
		private static readonly RichPresence Presence = Shared.Shared.GetNewPresence("excel");

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			this.Client = new DiscordRpcClient(Shared.Shared.GetString("discordID"));
			this.Client.Initialize();
			this.Client.SetPresence(Presence);

			this.Application.WorkbookDeactivate += this.Application_WorkbookDeactivate;
			this.Application.WorkbookOpen += this.Application_WorkbookOpen;
			((AppEvents_Event)this.Application).NewWorkbook += this.Application_WorkbookOpen;
		}

		private void Application_WorkbookOpen(Workbook wb)
		{
			Presence.Details = this.Application.ActiveWorkbook.Name;
			Presence.State = Shared.Shared.GetString("editing");
			Presence.Assets.LargeImageKey = "excel_editing";

			this.Client.SetPresence(Presence);
		}

		private void Application_WorkbookDeactivate(Workbook wb)
		{
			if (this.Application.Workbooks.Count == 1)
			{
				Presence.Details = Shared.Shared.GetString("noFile");
				Presence.State = null;
				Presence.Assets.LargeImageKey = "excel_nothing";
			}
			else
			{
				Presence.Details = this.Application.ActiveWorkbook.Name;
			}

			this.Client.SetPresence(Presence);
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			this.Client.Dispose();
		}

#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += this.ThisAddIn_Startup;
			this.Shutdown += this.ThisAddIn_Shutdown;
		}

#endregion
	}
}