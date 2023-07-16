using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemOfficeSeparator : WorkspaceItemBase
    {
        private const string Coma = "[,]";

        public override WorkspaceItem ItemType => WorkspaceItem.ExcelCultureFormat;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.Middle;
        
        public override string Help =>
            $"Определение настроенного в Windows разделителя дробной части числа{Environment.NewLine}" +
            "[Excel] >> [Файл] >> [Параметры] >> " +
            $"[Дополнительно] >> [Разделитель целой и дробной части]{Environment.NewLine}" +
            $"Необходимо указать символ: '{Coma}' (запятая),{Environment.NewLine}" +
            "и установить настройку \"Использовать системные разделители\"";

        public override string[] OptimalValues => new[] { Coma };

        public override bool ActualIsOptimal => ActualValue.Equals(Coma);

        public override bool UpdateStateAction(out string actualValue)
        {
            actualValue = string.Empty;

            Application application = null;

            try
            {
                application = new Application();
                actualValue = $"[{application.DecimalSeparator}]";
            }
            finally
            {
                if (application != null)
                {
                    int cnt = Marshal.ReleaseComObject(application);
                    if (cnt > 0)
                    {
                        Debug.WriteLine($"DEBUG => RELEASE COM [{nameof(application)}] Links.Count = {cnt}");
                    }
                    // ReSharper disable once RedundantAssignment
                    application = null;
                }
            }

            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
