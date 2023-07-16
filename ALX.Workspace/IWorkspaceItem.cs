using System.Drawing;
using System.ComponentModel;

namespace ALX.Workspace
{
    /// <summary>
    /// Приоритет элемента рабочего окружения
    /// </summary>
    public enum WorkspaceItemPriority
    {
        /// <summary>
        /// Низкий
        /// </summary>
        [Description("Низкий")]
        [EnumMemberColor(nameof(Color.MediumSeaGreen))]
        Low = 0,
        /// <summary>
        /// Средний
        /// </summary>
        [Description("Средний")]
        [EnumMemberColor(nameof(Color.Orange))]
        Middle = 1,
        /// <summary>
        /// Высокий
        /// </summary>
        [Description("Высокий")]
        [EnumMemberColor(nameof(Color.OrangeRed))]
        High = 2
    }

    /// <summary>
    /// Состояние элемента рабочего окружения
    /// </summary>
    public enum WorkspaceItemState
    {
        /// <summary>
        /// Нет данных
        /// </summary>
        [Description("Нет данных")]
        NoData = 0,
        /// <summary>
        /// Ожидание
        /// </summary>
        [Description("Ожидание")]
        Wait = 1,
        /// <summary>
        /// Актуально
        /// </summary>
        [Description("Актуально")]
        Updated = 2
    }

    /// <summary>
    /// Элемент рабочего окружения
    /// </summary>
    public enum WorkspaceItem
    {
        /// <summary>
        /// Пользовательский
        /// </summary>
        [Description("Пользовательский")]
        Custom = 0,
        /// <summary>
        /// Операционная система
        /// </summary>
        [Description("Операционная система")]
        WindowsVersion = 1,
        /// <summary>
        /// Microsoft Office
        /// </summary>
        [Description("Microsoft Office")]
        OfficeVersion = 2,
        /// <summary>
        /// Активация Microsoft Office
        /// </summary>
        [Description("Активация Microsoft Office")]
        OfficeActivation = 3,
        /// <summary>
        /// .NET Framework 4.7.2
        /// </summary>
        [Description(".NET Framework 4.7.2")]
        Framework472 = 4,
        /// <summary>
        /// .NET Framework
        /// </summary>
        [Description(".NET Framework")]
        FrameworkVersion = 5,
        /// <summary>
        /// Visual Studio Tools for Office 2010
        /// </summary>
        [Description("Visual Studio Tools for Office 2010")]
        Vsto2010 = 6,
        /// <summary>
        /// Русская раскладка (ru-RU)
        /// </summary>
        [Description("Русская раскладка (ru-RU)")]
        RuCulture = 7,
        /// <summary>
        /// Разделитель целой и дробной части числа в Windows
        /// </summary>
        [Description("Разделитель целой и дробной части числа в Windows")]
        WindowsCultureFormat = 8,
        /// <summary>
        /// Разделитель целой и дробной части числа в Microsoft Excel
        /// </summary>
        [Description("Разделитель целой и дробной части числа в Excel")]
        ExcelCultureFormat = 9,
        /// <summary>
        /// Наличие корневые сертификаты
        /// </summary>
        [Description("Корневой сертификат")]
        RootCerts = 10,
        /// <summary>
        /// Наличие старых версий Microsoft Office
        /// </summary>
        [Description("Наличие старых версий Microsoft Office")]
        OfficeOldVersions = 11,
        /// <summary>
        /// Переносимый профиль пользователя
        /// </summary>
        [Description("Переносимый профиль пользователя")]
        PortableUserProfile = 12,
        /// <summary>
        /// .NET Framework 4.8
        /// </summary>
        [Description(".NET Framework 4.8")]
        Framework48 = 13,
    }

    /// <summary>
    /// Интерфейс элемент рабочего окружения
    /// </summary>
    public interface IWorkspaceItem
    {
        /// <summary>
        /// Тип элемента рабочего окружения
        /// </summary>
        WorkspaceItem ItemType { get; }

        /// <summary>
        /// Имя для отображения
        /// </summary>
        string DisplayName { get; }

        /// <summary>
        /// Приоритет
        /// </summary>
        WorkspaceItemPriority Priority { get; }

        /// <summary>
        /// Справка/рекомендации
        /// </summary>
        string Help { get; }

        /// <summary>
        /// Оптимальные значение
        /// </summary>
        string[] OptimalValues { get; }

        /// <summary>
        /// Состояние
        /// </summary>
        WorkspaceItemState State { get; }

        /// <summary>
        /// Описание текущего состояния
        /// </summary>
        string StateDetails { get; }

        /// <summary>
        /// Признак того, что актуальное = оптимальному
        /// </summary>
        bool ActualIsOptimal { get; }

        /// <summary>
        /// Актуальное значение
        /// </summary>
        string ActualValue { get; }
    }
}