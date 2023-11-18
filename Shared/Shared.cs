using DiscordRPC;

using System.Collections.Generic;
using System.Diagnostics;

namespace Shared
{
	public class Shared
	{
		private static readonly IDictionary<int, string> OfficeVersions = new Dictionary<int, string>()
		{
			{ 6, "4.x" },
			{ 7, "95" },
			{ 8, "97" },
			{ 9, "2000" },
			{ 10, "XP" },
			{ 11, "2003" },
			{ 12, "2007" },
			{ 14, "2010" },
			{ 15, "2013" },
			{ 16, "2016 or 2019" }
		};

		private static readonly IDictionary<string, string> Strings = new Dictionary<string, string>()
		{
			{ "discordID", "470239659591598091" },
			{ "discordVisioID", "1175416723432800326" },
			{ "noFile", "No File Open" },
			{ "tabOut", "Not Active" },
			{ "welcome", "Welcome Screen" },
			{ "editing", "Editing File" },
			{ "editingSlide", "Editing Slide" },
			{ "editingPage", "Editing Page" },
			{ "editingDiagram", "Editing Diagram" },
			{ "presenting", "Presenting" },
			{ "excel", "Microsoft Excel" },
			{ "powerpoint", "Microsoft PowerPoint" },
			{ "word", "Microsoft Word" },
			{ "outlook", "Microsoft Outlook" },
			{ "unknown_version", "[Unknown Version]" },
			{ "unknown_key", "[Unknown]" },
			{ "visio", "Microsoft Visio" }
		};

		public static string GetVersion()
		{
			var version = Process.GetCurrentProcess().MainModule.FileVersionInfo.ProductMajorPart;
			return OfficeVersions.TryGetValue(version, out var versionValue) ? versionValue : GetString("unknown_version");
		}

		public static string GetString(string key)
			=> Strings.TryGetValue(key, out var s) ? s : GetString("unknown_key");

		public static RichPresence GetNewPresence(string type)
			=> new RichPresence()
			{
				Details = GetString("noFile"),
				State = GetString("welcome"),
				Assets = new Assets()
				{
					LargeImageKey = type + "_welcome",
					LargeImageText = GetString(type) + " " + GetVersion(),
					SmallImageKey = type
				}
			};
	}
}