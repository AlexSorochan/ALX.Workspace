using System;
using System.Threading;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemWindowsSeparator : WorkspaceItemBase
    {
        private const string Coma = "[,]";

        public override WorkspaceItem ItemType => WorkspaceItem.WindowsCultureFormat;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.High;
        
        public override string Help => 
            $"Определение настроенного в Windows разделителя дробной части числа{Environment.NewLine}" +
            "[Меню \"Пуск\"] >> [Панель управления] >> [Изменение форматов даты, времени и числа] >> " +
            $"[Дополнительные параметры] >> [Разделитель целой и дробной части]{Environment.NewLine}" +
            $"Необходимо указать символ: '{Coma}' (запятая)";

        public override string[] OptimalValues => new[] { Coma };

        public override bool ActualIsOptimal => ActualValue.Equals(Coma);

        public override bool UpdateStateAction(out string actualValue)
        {
            string value = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            actualValue = $"[{value}]";
            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
