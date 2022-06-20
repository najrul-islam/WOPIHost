using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WopiHost.Discovery.Enumerations;

namespace WopiHost.Discovery
{
	public interface IDiscoverer
	{
		Task<string> GetUrlTemplateAsync(string extension, WopiActionEnum action);
		Task<bool> SupportsExtensionAsync(string extension);
		Task<bool> SupportsActionAsync(string extension, WopiActionEnum action);
		Task<IEnumerable<string>> GetActionRequirementsAsync(string extension, WopiActionEnum action);
		Task<bool> RequiresCobaltAsync(string extension, WopiActionEnum action);
		Task<string> GetApplicationNameAsync(string extension);
		/// <summary>
		/// Gets the icon of the application that handles files with the given extension.
		/// </summary>
		/// <param name="extension">File extension to get the icon for (without the leading dot).</param>
		/// <returns>Icon of the app.</returns>
		Task<Uri> GetApplicationFavIconAsync(string extension);
		Task<WopiProof> GetWopiProof();
	}
}