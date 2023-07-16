using System;
using Microsoft.Win32;

namespace ALX.Workspace.Console
{
    public class CustomWorkspaceItem : WorkspaceItemBase
    {
        public const string
            InstalledStr = "Установлен",
            NotInstalledStr = "Не установлен",
            RegParamDisplayName = "DisplayName",
            AppName = "APB.SV.POSTerminal.Arcus2",
            RegKey64 = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
            RegKey32 = "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall";

        public override WorkspaceItem ItemType => WorkspaceItem.Custom;

        public override string DisplayName => "Установка Arcus2";

        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.High;

        public override string Help => "При установке Arcus2, не рекомендуется менять каталог установки по умолчанию.\n\r" +
            "После успешной становки, в корне диска C:\\ должен быть каталог \"Arcus2\"\n\r" +
            "В случае его отсутствия, рекомендуется переустановить Arcus2";

        public override string[] OptimalValues => new[] { InstalledStr, NotInstalledStr };

        public override bool ActualIsOptimal => string.Equals(ActualValue, InstalledStr);

        public override bool UpdateStateAction(out string actualValue)
        {
            bool installed = false;
            actualValue = string.Empty;

            try
            {
                // search in: CurrentUser
                installed = CheckInstallationInRegKey(inCurrentUser: true, subKeyName: RegKey64);
                // search in: LocalMachine_32
                if (!installed) installed = CheckInstallationInRegKey(inCurrentUser:false, subKeyName: RegKey64);
                // search in: LocalMachine_64
                if (!installed) installed = CheckInstallationInRegKey(inCurrentUser: false, subKeyName: RegKey32);
            }
            catch (Exception)
            {
                installed = false;
            }

            // NOT FOUND

            actualValue = installed ? InstalledStr : NotInstalledStr;
            return !string.IsNullOrEmpty(actualValue);
        }

        private bool CheckInstallationInRegKey(bool inCurrentUser, string subKeyName)
        {
            bool installed = false;
            string displayName = string.Empty;

            RegistryKey
                parentNode = inCurrentUser ? Registry.CurrentUser : Registry.LocalMachine,
                key = parentNode.OpenSubKey(subKeyName),
                subkey = null;

            if (key != null)
                foreach (string keyName in key.GetSubKeyNames())
                {
                    subkey = key.OpenSubKey(keyName);
                    if (subkey == null) continue;
                    displayName = subkey.GetValue(RegParamDisplayName)?.ToString() ?? string.Empty;
                    installed = AppName.Equals(displayName, StringComparison.OrdinalIgnoreCase);
                    if (installed) break;
                }

            return installed;
        }
    }
}
