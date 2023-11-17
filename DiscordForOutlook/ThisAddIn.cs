using DiscordRPC;

using Microsoft.Office.Interop.Outlook;

using System;

namespace DiscordForOutlook
{
	public partial class ThisAddIn
	{
		public DiscordRpcClient Client;
		private static readonly RichPresence Presence = Shared.Shared.GetNewPresence("outlook");

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			this.Client = new DiscordRpcClient(Shared.Shared.GetString("discordID"));
			this.Client.Initialize();
			Presence.State = null;
			Presence.Details = null;
			Presence.Assets.LargeImageKey = "outlook_info";
			this.Client.SetPresence(Presence);

			((ApplicationEvents_11_Event)this.Application).Quit += this.ThisAddIn_Quit;
		}

		private void ThisAddIn_Quit()
		{
			this.Client.Dispose();
			return;
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			// Note: Outlook no longer raises this event. If you have code that 
			//    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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