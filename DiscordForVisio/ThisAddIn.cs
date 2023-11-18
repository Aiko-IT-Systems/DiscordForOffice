using DiscordRPC;

using Microsoft.Office.Interop.Visio;

using System;

namespace DiscordForVisio
{
	public partial class ThisAddIn
	{
		public DiscordRpcClient Client;
		private static readonly RichPresence Presence = Shared.Shared.GetNewPresence("visio");

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			this.Client = new DiscordRpcClient(Shared.Shared.GetString("discordVisioID"));
			this.Client.Initialize();
			this.Client.SetPresence(Presence);

			this.Application.WindowActivated += this.ApplicationOnWindowActivated;
			this.Application.DocumentOpened += this.ApplicationOnDocumentOpened;
			this.Application.DocumentChanged += this.ApplicationOnDocumentChanged;
			this.Application.BeforeWindowClosed += this.Application_WindowDeactivate;

			try
			{
				// Use the currently opened document
				var doc = this.Application.ActiveDocument;
				this.Application_DocumentOpen(doc);
			}
			catch
			{
				// Use the default presence when there is no current document
			}
		}

		private void SetPresence()
		{
			Presence.Details = this.Application.ActiveDocument.Name;
			Presence.State = Shared.Shared.GetString("editingDiagram");
			Presence.Assets.LargeImageKey = "visio_editing";

			this.Client.SetPresence(Presence);
		}

		private void Application_WindowDeactivate(Window wn)
		{
			Presence.Details = Shared.Shared.GetString("tabOut");
			Presence.State = null;
			Presence.Assets.LargeImageKey = "visio_nothing";

			this.Client.SetPresence(Presence);
		}

		private void Application_WindowClose(Document doc)
		{
			Presence.Details = Shared.Shared.GetString("tabOut") + this.Application.Documents.Count;
			Presence.State = null;
			Presence.Assets.LargeImageKey = "visio_nothing";

			this.Client.SetPresence(Presence);
		}

		private void Application_DocumentOpen(Document doc)
		{
			this.SetPresence();
			doc.BeforeDocumentClose += this.Application_WindowClose;
		}

		private void ApplicationOnDocumentChanged(Document doc)
		{
			if (this.Application.Documents.Count == 1)
			{
				this.SetPresence();
			}
		}

		private void ApplicationOnDocumentOpened(Document doc)
		{
			this.SetPresence();
		}

		private void ApplicationOnWindowActivated(Window window)
		{
			this.SetPresence();
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