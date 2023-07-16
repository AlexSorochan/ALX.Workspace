using System;
using Microsoft.Win32;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemVsto2010 : WorkspaceItemBase
    {
        private const string
            InstalledStr = "Установлен",
            NotInstalledStr = "Не установлен",
            RegParamDisplayName = "DisplayName",
            Key32 = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
            AppName32 = "Microsoft Visual Studio 2010 Tools for Office Runtime (x86)",
            AppName64 = "Microsoft Visual Studio 2010 Tools for Office Runtime (x64)";

        public override WorkspaceItem ItemType => WorkspaceItem.Vsto2010;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.High;
        
        public override string Help =>
            $"Поиск установленного приложения: [{(Environment.Is64BitOperatingSystem ? AppName64 : AppName32)}]" +
            $"{Environment.NewLine}По ключу реестра:{Environment.NewLine}" +
            $"[Компьютер\\HKEY_LOCAL_MACHINE\\{Key32}]";

        public override string[] OptimalValues => new[] { InstalledStr, NotInstalledStr };

        public override bool ActualIsOptimal => string.Equals(ActualValue, InstalledStr);

        public override bool UpdateStateAction(out string actualValue)
        {
            bool installed = false;
            actualValue = NotInstalledStr;

            string appName = Environment.Is64BitOperatingSystem
                ? AppName64 : AppName32;

            using (RegistryKey keyLocalMachine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32))
            using (RegistryKey keyUninstall = keyLocalMachine.OpenSubKey(Key32))
            {
                if (keyUninstall == null) return false;
                string[] subKeyNames = keyUninstall.GetSubKeyNames();

                foreach (string keyName in subKeyNames)
                {
                    // ReSharper disable once AccessToDisposedClosure
                    using (RegistryKey subKey = keyUninstall.OpenSubKey(keyName))
                    {
                        string displayName = subKey?.GetValue(RegParamDisplayName)?.ToString() ?? string.Empty;
                        installed = appName.Equals(displayName, StringComparison.OrdinalIgnoreCase);
                        if (installed) break;
                    }
                }
            }

            actualValue = installed ? InstalledStr : NotInstalledStr;

            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
