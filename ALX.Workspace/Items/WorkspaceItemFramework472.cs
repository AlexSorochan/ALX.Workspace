using System;
using Microsoft.Win32;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemFramework472 : WorkspaceItemBase
    {
        public const string
            InstalledStr = "Установлен",
            NotInstalledStr = "Не установлен",
            RegParamDisplayName = "DisplayName",
            AppName = "Microsoft .NET Framework 4.7.2 Targeting Pack",
            Key32 = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
            Key64 = "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall";

        public override WorkspaceItem ItemType => WorkspaceItem.Framework472;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.Middle;
        
        public override string Help => 
            $"Поиск установленного приложения: [{AppName}]" +
            $"{Environment.NewLine}По ключу реестра:{Environment.NewLine}" +
            $"[Компьютер\\HKEY_LOCAL_MACHINE\\{Key32}]";

        public override string[] OptimalValues => new[] { InstalledStr, NotInstalledStr };

        public override bool ActualIsOptimal => string.Equals(ActualValue, InstalledStr);

        public override bool UpdateStateAction(out string actualValue)
        {
            bool installed = false;
            actualValue = string.Empty;

            using (RegistryKey keyLocalMachine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32))
            {
                using (RegistryKey keyUninstall = keyLocalMachine.OpenSubKey(
                    Environment.Is64BitOperatingSystem ? Key64 : Key32))
                {
                    if (keyUninstall == null) return false;
                    string[] subKeyNames = keyUninstall.GetSubKeyNames();

                    foreach (string keyName in subKeyNames)
                    {
                        // ReSharper disable once AccessToDisposedClosure
                        using (RegistryKey subKey = keyUninstall.OpenSubKey(keyName))
                        {
                            string displayName = subKey?.GetValue(RegParamDisplayName)?.ToString() ?? string.Empty;
                            installed = AppName.Equals(displayName, StringComparison.OrdinalIgnoreCase);
                            if (installed) break;
                        }
                    }
                }
            }

            actualValue = installed ? InstalledStr : NotInstalledStr;
            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
