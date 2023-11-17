using DiscordRPC;

using Microsoft.Office.Interop.Word;

using System;

namespace DiscordForWord
{
	public partial class ThisAddIn
	{
		public DiscordRpcClient Client;
		private static readonly RichPresence Presence = Shared.Shared.GetNewPresence("word");

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			this.Client = new DiscordRpcClient(Shared.Shared.GetString("discordID"));
			this.Client.Initialize();
			this.Client.SetPresence(Presence);

			this.Application.WindowDeactivate += this.Application_WindowDeactivate;
			this.Application.WindowActivate += this.Application_WindowActivate;
			this.Application.DocumentOpen += this.Application_DocumentOpen;
			((ApplicationEvents4_Event)this.Application).NewDocument += this.Application_DocumentOpen;
			this.Application.WindowSelectionChange += this.Application_WindowSelectionChange;
			this.Application.DocumentChange += this.Application_DocumentChange;

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

		private void Application_DocumentChange()
		{
			if (this.Application.Documents.Count == 1)
			{
				this.Application_WindowSelectionChange(this.Application.Selection);
			}
		}

		private void Application_WindowDeactivate(Document doc, Window wn)
		{
			if (this.Application.Documents.Count == 1)
			{
				Presence.Details = Shared.Shared.GetString("tabOut");
				Presence.State = null;
				Presence.Party = null;
				Presence.Assets.LargeImageKey = "word_nothing";
			}

			this.Client.SetPresence(Presence);
		}

		private void Application_WindowClose()
		{
			if (this.Application.Documents.Count > 1)
			{
				Presence.Details = "" + this.Application.Documents.Count;
				this.Application_WindowSelectionChange(this.Application.Selection);
			}
			else
			{
				Presence.Details = Shared.Shared.GetString("tabOut") + this.Application.Documents.Count;
				Presence.State = null;
				Presence.Party = null;
				Presence.Assets.LargeImageKey = "word_nothing";
			}

			this.Client.SetPresence(Presence);
		}

		private void Application_DocumentOpen(Document doc)
		{
			this.Application_WindowSelectionChange(this.Application.Selection);

			((DocumentEvents2_Event)doc).Close += this.Application_WindowClose;
		}

		private void Application_WindowActivate(Document doc, Window wn)
		{
			this.Application_WindowSelectionChange(this.Application.Selection);
		}

		public void Application_WindowSelectionChange(Selection sel)
		{
			var range = this.Application.ActiveDocument.Content;

			Presence.Details = this.Application.ActiveDocument.Name;
			Presence.State = Shared.Shared.GetString("editingPage");
			Presence.Assets.LargeImageKey = "word_editing";
			Presence.Party = new Party()
			{
				ID = Secrets.CreateFriendlySecret(new Random()),
				Max = range.ComputeStatistics(WdStatistic.wdStatisticPages),
				Size = (int)sel.Information[WdInformation.wdActiveEndPageNumber]
			};

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