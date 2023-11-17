using DiscordRPC;

using Microsoft.Office.Interop.PowerPoint;

using System;

namespace DiscordForPowerPoint
{
	public partial class ThisAddIn
	{
		public DiscordRpcClient Client;
		private static readonly RichPresence Presence = Shared.Shared.GetNewPresence("powerpoint");

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			this.Client = new DiscordRpcClient(Shared.Shared.GetString("discordID"));
			this.Client.Initialize();
			this.Client.SetPresence(Presence);

			// An event handler for when a new slide is created
			this.Application.PresentationNewSlide += this.Application_PresentationNewSlide;

			// An event handler for any time a slide / slides / inbetween slides is selected
			this.Application.SlideSelectionChanged += this.Application_SlideSelectionChanged;

			// An event handler for when a file is closed.
			// Final = Actually closed
			this.Application.PresentationCloseFinal += this.Application_PresentationCloseFinal;

			// Event handlers for when a file is created, opened, saved, or slide show ends.
			this.Application.AfterNewPresentation += this.Application_AfterPresentationOpenEvent;
			this.Application.AfterPresentationOpen += this.Application_AfterPresentationOpenEvent;
			this.Application.PresentationSave += this.Application_AfterPresentationOpenEvent;
			this.Application.SlideShowEnd += this.Application_AfterPresentationOpenEvent;

			// An event handler for when a slide show starts, or goes onto a new slide
			this.Application.SlideShowNextSlide += this.Application_SlideShowNextSlide;
		}

		// When Microsoft PowerPoint shuts down, delete the RPC client.
		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			this.Client.Dispose();
		}

		private void Application_PresentationNewSlide(Slide sld)
		{
			// Assumption: People start on slide 1 when creating a file
			Presence.Party = new Party()
			{
				Max = this.Application.ActivePresentation.Slides.Count,
				Size = 1
			};

			this.Client.SetPresence(Presence);
		}

		private void Application_SlideSelectionChanged(SlideRange sldRange)
		{
			if (sldRange.Count <= 0)
				return;

			Presence.Details = sldRange.Application.ActivePresentation.Name;
			Presence.State = Shared.Shared.GetString("editing");
			Presence.Assets.LargeImageKey = "editing";
			Presence.Party = new Party()
			{
				ID = Secrets.CreateFriendlySecret(new Random()),
				Max = this.Application.ActivePresentation.Slides.Count,
				Size = sldRange[1].SlideNumber
			};
			this.Client.SetPresence(Presence);
		}

		public void Application_PresentationCloseFinal(Presentation pres)
		{
			// There's only one presentation left - the current one
			if (this.Application.Presentations.Count == 1)
			{
				Presence.Details = Shared.Shared.GetString("noFile");
				Presence.State = null;
				Presence.Party = null;
				Presence.Assets.LargeImageKey = "nothing";
			}
			else
			{
				Presence.Details = this.Application.ActivePresentation.Name;
			}

			this.Client.SetPresence(Presence);
		}

		public void Application_AfterPresentationOpenEvent(Presentation pres)
		{
			Presence.Details = pres.Name;
			Presence.State = Shared.Shared.GetString("editingSlide");
			Presence.Assets.LargeImageKey = "editing";

			// Slide selection is also triggered - Don't need to set presence
		}

		public void Application_SlideShowNextSlide(SlideShowWindow wn)
		{
			Presence.Details = wn.Presentation.Name;
			Presence.State = Shared.Shared.GetString("presenting");
			Presence.Assets.LargeImageKey = "present";
			Presence.Party = new Party()
			{
				ID = Secrets.CreateFriendlySecret(new Random()),
				Max = wn.Presentation.Slides.Count,
				Size = wn.View.CurrentShowPosition
			};
			this.Client.SetPresence(Presence);
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