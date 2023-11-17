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
			this.Client = new DiscordRpcClient(Shared.Shared.GetString("discordID"));
			this.Client.Initialize();
			this.Client.SetPresence(Presence);

			this.Application.WindowActivated += this.ApplicationOnWindowActivated;
			this.Application.DocumentOpened += this.ApplicationOnDocumentOpened;
			this.Application.SelectionChanged += this.ApplicationOnSelectionChanged;
			this.Application.DocumentChanged += this.ApplicationOnDocumentChanged;

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

		private void Application_DocumentOpen(Document doc)
		{
			throw new NotImplementedException();
		}

		private void ApplicationOnDocumentChanged(Document doc)
		{
			throw new NotImplementedException();
		}

		private void ApplicationOnSelectionChanged(Window window)
		{
			throw new NotImplementedException();
		}

		private void ApplicationOnDocumentOpened(Document doc)
		{
			throw new NotImplementedException();
		}

		private void ApplicationOnWindowActivated(Window window)
		{
			throw new NotImplementedException();
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