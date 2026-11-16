// ApiKeys.cs (PowerPoint - C# 7.3 / .NET 4.8, VSTO)
// User-supplied API keys for the online features: logo search (Brave Search),
// icon search (Freepik) and map geocoding (Mapbox).

using Microsoft.Win32;
using System;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace OfficePragmatics_PowerPoint.Settings
{
    /// <summary>
    /// Each user brings their own keys for the third-party services the add-in calls.
    /// Keys are encrypted with Windows DPAPI for the current Windows user and stored in
    /// HKCU, separate from settings.json, so they are never written to disk in plain text.
    /// </summary>
    public static class ApiKeys
    {
        private const string RegistryKeyPath = @"Software\Office Pragmatics\ApiKeys";
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("OfficePragmatics.ApiKeys.v1");

        public const string BraveSearchApiKeyUrl = "https://brave.com/search/api/";
        public const string FreepikApiKeyUrl = "https://www.freepik.com/api";
        public const string MapboxTokenUrl = "https://account.mapbox.com/access-tokens/";

        /// <summary>Brave Search API subscription token (logo search via image search).</summary>
        public static string BraveSearchApiKey
        {
            get => Get("BraveSearchApiKey");
            set => Set("BraveSearchApiKey", value);
        }

        /// <summary>Freepik API key (icon search and SVG download).</summary>
        public static string FreepikApiKey
        {
            get => Get("FreepikApiKey");
            set => Set("FreepikApiKey", value);
        }

        /// <summary>Mapbox access token (geocoding for Plot Map).</summary>
        public static string MapboxAccessToken
        {
            get => Get("MapboxAccessToken");
            set => Set("MapboxAccessToken", value);
        }

        public static bool HasLogoSearchKey => !string.IsNullOrEmpty(BraveSearchApiKey);

        public static bool HasIconSearchKey => !string.IsNullOrEmpty(FreepikApiKey);

        public static bool HasMapsKey => !string.IsNullOrEmpty(MapboxAccessToken);

        // ---------------------------------------------------------
        // Feature gates: call at the start of a feature. If its key is
        // missing, offer to open Settings on the API Keys page.
        // ---------------------------------------------------------

        public static bool EnsureLogoSearchKey() =>
            EnsureConfigured(() => HasLogoSearchKey,
                "Logo search uses Brave Search with your own Brave Search API key.");

        public static bool EnsureIconSearchKey() =>
            EnsureConfigured(() => HasIconSearchKey,
                "Icon search uses the Freepik API with your own Freepik API key.");

        public static bool EnsureMapsKey() =>
            EnsureConfigured(() => HasMapsKey,
                "Plot Map uses Mapbox geocoding with your own Mapbox access token.");

        private static bool EnsureConfigured(Func<bool> isConfigured, string explanation)
        {
            if (isConfigured()) return true;

            var answer = MessageBox.Show(
                explanation + "\n\nOpen Settings to add it now?",
                "API Key Needed",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (answer != DialogResult.Yes) return false;

            SettingsForm.ShowSettingsDialog(SettingsForm.ApiKeysPageName);
            return isConfigured();
        }

        // ---------------------------------------------------------
        // Storage
        // ---------------------------------------------------------

        private static string Get(string name)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key?.GetValue(name) is byte[] protectedBytes)
                    {
                        byte[] plain = ProtectedData.Unprotect(protectedBytes, Entropy, DataProtectionScope.CurrentUser);
                        return Encoding.UTF8.GetString(plain);
                    }
                }
            }
            catch
            {
                // Unreadable (e.g. copied from another Windows account): treat as not set.
            }
            return string.Empty;
        }

        private static void Set(string name, string value)
        {
            value = (value ?? string.Empty).Trim();

            using (var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                if (value.Length == 0)
                {
                    key.DeleteValue(name, throwOnMissingValue: false);
                    return;
                }

                byte[] protectedBytes = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(value), Entropy, DataProtectionScope.CurrentUser);
                key.SetValue(name, protectedBytes, RegistryValueKind.Binary);
            }
        }
    }
}
