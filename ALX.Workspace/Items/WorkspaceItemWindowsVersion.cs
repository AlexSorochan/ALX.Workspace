using System;
using System.Linq;
using Microsoft.Win32;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemWindowsVersion : WorkspaceItemBase
    {
        private const string
            KeyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
            ValueName = "ProductName",
            Win7 = "Windows 7";

        public override WorkspaceItem ItemType => WorkspaceItem.WindowsVersion;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.High;

        public override string Help => 
            $"Ключ реестра: [Компьютер\\HKEY_LOCAL_MACHINE\\{KeyPath}]" +
            $"{Environment.NewLine}Имя параметра: [{ValueName}]";

        public override string[] OptimalValues => new[]
        {
            "Windows 7 Starter",
            "Windows 7 Home Basic",
            "Windows 7 Home Premium",
            "Windows 7 Professional",
            "Windows 7 Enterprise",
            "Windows 7 Ultimate",
            "Windows 8",
            "Windows 8 Pro",
            "Windows 8 Enterprise",
            "Windows 8.1",
            "Windows 8.1 Pro",
            "Windows 8.1 Enterprise",
            "Windows 10 Home",
            "Windows 10 Pro",
            "Windows 10 Enterprise",
            "Windows 10 Education",
            "Windows 10 Pro Education",
            "Windows 10 Enterprise LTSB",
            "Windows 10 Enterprise LTSC",
            "Windows 10 Pro for Workstations",
            "Windows 11 Home",
            "Windows 11 Pro",
            "Windows 11 Pro for Workstations",
            "Windows 11 Pro Education",
            "Windows 11 Enterprise",
            "Windows 11 Education"
        };

        public override bool ActualIsOptimal => OptimalValues.Any(x => ActualValue.StartsWith(x));

        public override bool UpdateStateAction(out string actualValue)
        {
            actualValue = string.Empty;

            using(RegistryKey regKey = Registry.LocalMachine.OpenSubKey(KeyPath))
            {
                if (regKey == null)
                    throw new Exception("Ключ реестра не найден");

                object keyValue = regKey.GetValue(ValueName);
                if (keyValue == null)
                    throw new Exception("Значение не найдено");

                string windowsVersion = keyValue.ToString();

                // todo проверить SP2 в Win7
                // https://docs.microsoft.com/en-us/dotnet/api/system.operatingsystem.servicepack?view=net-6.0

                actualValue = 
                    $"{windowsVersion} " +
                    $"(x{(Environment.Is64BitOperatingSystem ? "64" : "32")})";
            }
            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
