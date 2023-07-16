using System;
using System.Linq;
using System.Runtime.Serialization;
using ALX.Workspace.Items;

namespace ALX.Workspace
{
    /// <summary>
    /// Базовый класс элемента рабочего окружения
    /// </summary>
    [KnownType(typeof(WorkspaceItemWindowsVersion))]
    [KnownType(typeof(WorkspaceItemFrameworkVersion))]
    [KnownType(typeof(WorkspaceItemFramework472))]
    [KnownType(typeof(WorkspaceItemFramework48))]
    [KnownType(typeof(WorkspaceItemVsto2010))]
    [KnownType(typeof(WorkspaceItemOfficeVersion))]
    [KnownType(typeof(WorkspaceItemOfficeActivated))]
    [KnownType(typeof(WorkspaceItemCultureRu))]
    [KnownType(typeof(WorkspaceItemWindowsSeparator))]
    [KnownType(typeof(WorkspaceItemOfficeSeparator))]
    [KnownType(typeof(WorkspaceItemRootCerts))]
    public abstract class WorkspaceItemBase : IWorkspaceItem
    {
        private const string 
            UpdatedDetails = "Актуальное состояние",
            WaitDetails = "Подождите ...";

        /// <summary>
        /// Тип элемента рабочего окружения
        /// </summary>
        public abstract WorkspaceItem ItemType { get; }

        /// <summary>
        /// Имя для отображения
        /// </summary>
        public abstract string DisplayName { get; }
        
        /// <summary>
        /// Имя для отображения (HTML)
        /// </summary>
        public string DisplayNameHtmlUi 
        {
            get
            {
                string imageName = Priority == WorkspaceItemPriority.High
                    ? ActualIsOptimal ? "ok_16" : "cancel_16"
                    : ActualIsOptimal ? "ok_16" : "warning_16";

                return $"<image={imageName}> {DisplayName}";
            }
        }

        /// <summary>
        /// Приоритет
        /// </summary>
        public abstract WorkspaceItemPriority Priority { get; }

        /// <summary>
        /// Справка/рекомендации
        /// </summary>
        public abstract string Help { get; }

        /// <summary>
        /// Оптимальные значение
        /// </summary>
        public abstract string[] OptimalValues { get; }

        /// <summary>
        /// Состояние
        /// </summary>
        public WorkspaceItemState State { get; private set; }

        /// <summary>
        /// Описание текущего состояния
        /// </summary>
        public string StateDetails { get; private set; }

        /// <summary>
        /// Признак того, что актуальное является оптимальным
        /// </summary>
        public virtual bool ActualIsOptimal => OptimalValues.Contains(ActualValue);

        /// <summary>
        /// Актуальное значение
        /// </summary>
        public string ActualValue { get; private set; }

        /// <summary>
        /// Метод вызывается в public UpdateState
        /// </summary>
        /// <param name="actualValue">Актуальное значение</param>
        /// <returns>Признак успешного обновления текущего состояния</returns>
        public abstract bool UpdateStateAction(out string actualValue);

        /// <summary>
        /// Обновить состояние
        /// </summary>
        /// <returns>Признак успешного обновления состояния</returns>
        public bool UpdateState()
        {
            bool complete;

            try
            {
                // Установить состояние "Ожидание"
                State = WorkspaceItemState.Wait;
                StateDetails = WaitDetails;

                // Обновить состояние
                complete = UpdateStateAction(out string actualValue);
                ActualValue = actualValue;

                // Установить состояние "Обновлено"
                State = WorkspaceItemState.Updated;
                StateDetails = UpdatedDetails;
            }
            catch (Exception exception)
            {
                // Установить состояние "Нет данных" + "Ошибка"
                State = WorkspaceItemState.NoData;
                StateDetails = exception.ToString();

                // Вернуть "Неудачу"
                complete = false;
            }

            return complete;
        }
        
        public override string ToString()
        {
            try
            {
                return $"[{DisplayName}] => Value: [{ActualValue}], is optimal: [{ActualIsOptimal}], Level: [{Priority}]";
            }
            catch
            {
                return base.ToString();
            }
        }
    }
}
