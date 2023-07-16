using System;
using System.Linq;
using System.Globalization;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemCultureRu : WorkspaceItemBase
    {
        private const string
            CultureCode = "ru-RU",
            Installed = "Установлена",
            NotInstalled = "Не установлена";

        public override WorkspaceItem ItemType => WorkspaceItem.RuCulture;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.Middle;
        
        public override string Help => 
            $"Проверка наличия установленного русского языка в системе{Environment.NewLine}" +
            $"Раскладка с кодом: [{CultureCode}]";

        public override string[] OptimalValues => new[] { Installed, NotInstalled };

        public override bool ActualIsOptimal => ActualValue.Equals(Installed);

        public override bool UpdateStateAction(out string actualValue)
        {
            actualValue = string.Empty;

            CultureInfo[] cultures = CultureInfo.GetCultures(CultureTypes.AllCultures & ~CultureTypes.NeutralCultures);
            CultureInfo cultureRu = cultures.FirstOrDefault(c => c.Name.Equals(CultureCode, StringComparison.OrdinalIgnoreCase));

            actualValue = cultureRu != null ? Installed : NotInstalled;

            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
