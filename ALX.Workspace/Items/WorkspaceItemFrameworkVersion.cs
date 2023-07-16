using System;
using System.Linq;
using Microsoft.Win32;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemFrameworkVersion : WorkspaceItemBase
    {
        private const string
            KeyPath = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full",
            ValueName = "Version";

        public override WorkspaceItem ItemType => WorkspaceItem.FrameworkVersion;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.High;
        
        public override string Help => 
            $"Ключ реестра: [Компьютер\\HKEY_LOCAL_MACHINE\\{KeyPath}]" +
            $"{Environment.NewLine}Имя параметра: [{ValueName}]";

        public override string[] OptimalValues => new[]
        {
            "4.8",
            "4.7.2"
        };

        public override bool ActualIsOptimal => OptimalValues.Any(x => ActualValue.StartsWith(x));

        public override bool UpdateStateAction(out string actualValue)
        {
            actualValue = string.Empty;

            using (RegistryKey keyLocalMachine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32))
            {
                using (RegistryKey regKey = keyLocalMachine.OpenSubKey(KeyPath))
                {
                    object keyValue = regKey?.GetValue(ValueName);
                    if (keyValue == null) return false;

                    actualValue = keyValue.ToString();
                }
            }

            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
