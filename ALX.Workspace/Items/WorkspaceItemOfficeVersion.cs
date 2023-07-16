using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemOfficeVersion : WorkspaceItemBase
    {
        private const string
            XlExe = "EXCEL.EXE",
            AppName = "Microsoft Office",
            NotInstalledStr = "Не установлен";

        public override WorkspaceItem ItemType => WorkspaceItem.OfficeVersion;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.High;
        
        public override string Help => 
            $"Версия установленного {AppName}{Environment.NewLine}Номер версии начинается с:{Environment.NewLine}" +
            $"15.0 => Office 2013{Environment.NewLine}16.0 => Office 2016 / 2019 / 2021";

        public override string[] OptimalValues => new[] { "15.0", "16.0"};

        public override bool ActualIsOptimal => OptimalValues.Any(x => ActualValue.StartsWith(x));

        public override bool UpdateStateAction(out string actualValue)
        {
            actualValue = NotInstalledStr;

            Application application = null;

            try
            {
                application = new Application();
                string xlPath = Path.Combine(application.Path, XlExe);

                // Почему именно 8 ? Просто потому-что нет другого адекватного решения
                bool isX64 = Marshal.SizeOf(application.HinstancePtr) == 8;

                FileVersionInfo info = FileVersionInfo.GetVersionInfo(xlPath);
                actualValue = $"{info.ProductVersion} (x{(isX64 ? "64" : "32")})";
            }
            catch (Exception exception)
            {
                throw new Exception(
                    message:$"Ошибка при проверке версии Microsoft Office.{Environment.NewLine}{exception.Message}",
                    innerException:exception);
            }
            finally
            {
                if (application != null)
                {
                    int cnt = Marshal.ReleaseComObject(application);
                    if (cnt > 0) Debug.WriteLine($"DEBUG => RELEASE COM [{nameof(application)}] Links.Count = {cnt}");
                    // ReSharper disable once RedundantAssignment
                    application = null;
                }
            }

            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
