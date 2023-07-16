using System;
using System.Linq;
using System.Xml.Linq;
using System.Reflection;
using System.Collections.Generic;
using ALX.Workspace.Items;

namespace ALX.Workspace
{
    /// <summary>
    /// Менеджер рабочего окружения (singleton)
    /// </summary>
    public class WorkspaceManager
    {
        public string Description => "Параметры рабочего окружения";

        private static readonly List<WorkspaceItemBase> ItemsList = new List<WorkspaceItemBase>();
        
        private static WorkspaceManager _workspaceManager;

        /// <summary>
        /// Менеджер рабочего окружения
        /// </summary>
        protected WorkspaceManager(
            bool checkWindowsVersion, bool checkFrameworkVersion, bool checkFramework472,
            bool checkFramework48, bool checkVsto2010, bool checkOfficeVersion,
            bool checkOfficeActivated, bool checkCultureRu, bool checkWindowsSeparator,
            bool checkOfficeSeparator, bool checkRootCerts)
        {
            #region ПОЛУЧИТЬ ВСЕ КОМПОНЕНТЫ РАБОЧЕГО ОКРУЖЕНИЯ ПО УМОЛЧАНИЮ

            IEnumerable<Type> types = Assembly.GetExecutingAssembly().GetTypes().
                Where(x => x.IsSubclassOf(typeof(WorkspaceItemBase)));

            foreach(Type itemType in types)
            {
                object instance = Activator.CreateInstance(itemType);
                if (instance is WorkspaceItemBase item)
                {
                    switch (item.ItemType)
                    {
                        case WorkspaceItem.WindowsVersion:
                            if (checkWindowsVersion) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.OfficeVersion:
                            if (checkOfficeVersion) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.OfficeActivation:
                            if (checkOfficeActivated) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.Framework472: 
                            if (checkFramework472) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.FrameworkVersion: 
                            if (checkFrameworkVersion) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.Vsto2010: 
                            if (checkVsto2010) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.RuCulture: 
                            if (checkCultureRu) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.WindowsCultureFormat: 
                            if (checkWindowsSeparator) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.ExcelCultureFormat: 
                            if (checkOfficeSeparator) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.RootCerts: 
                            if (checkRootCerts) ItemsList.Add(item);
                            break;
                        case WorkspaceItem.Framework48: 
                            if (checkFramework48) ItemsList.Add(item);
                            break;
                    }
                }    
            }

            #endregion
        }

        /// <summary>
        /// Инициализация
        /// </summary>
        /// <param name="сheckWindowsVersion">Признак проверки версии ОС Windows</param>
        /// <param name="сheckFrameworkVersion">Признак проверки версии .NET Framework</param>
        /// <param name="сheckFramework472">Признак проверки наличия .NET Framework 4.7.2</param>
        /// <param name="сheckFramework48">Признак проверки наличия .NET Framework 4.7.2</param>
        /// <param name="сheckVsto2010">Признак проверки наличия Visual Studio Tool For Office 2010</param>
        /// <param name="сheckOfficeVersion">Признак проверки версии Microsoft Office</param>
        /// <param name="сheckOfficeActivated">Признак проверки активации Microsoft Office</param>
        /// <param name="сheckCultureRu">Признак проверки наличия русско раскладки</param>
        /// <param name="сheckWindowsSeparator">Признак проверки корректности децимального разделителя в ОС</param>
        /// <param name="сheckOfficeSeparator">Признак проверки корректности децимального разделителя в Office</param>
        /// <param name="сheckRootCerts">Признак проверки наличия корневых сертификатов</param>
        /// <returns></returns>
        public static WorkspaceManager Initialize(
            bool checkWindowsVersion = false, bool checkFrameworkVersion = false, bool checkFramework472 = false, 
            bool checkFramework48 = false, bool checkVsto2010 = false, bool checkOfficeVersion = false, 
            bool checkOfficeActivated = false, bool checkCultureRu = false, bool checkWindowsSeparator = false, 
            bool checkOfficeSeparator = false, bool checkRootCerts = false)
        {
            return _workspaceManager ?? (_workspaceManager = new WorkspaceManager(
                checkWindowsVersion, checkFrameworkVersion, checkFramework472,
                checkFramework48, checkVsto2010, checkOfficeVersion,
                checkOfficeActivated, checkCultureRu, checkWindowsSeparator,
                checkOfficeSeparator, checkRootCerts));
        }

        /// <summary>
        /// Элементы рабочего окружения
        /// </summary>
        public WorkspaceItemBase[] WorkspaceItems => ItemsList.OrderBy(x => x.ItemType).ToArray();

        /// <summary>
        /// Обновить значения параметров рабочего окружения
        /// </summary>
        public void Update()
        {
            ItemsList.ForEach(item => item.UpdateState());
        }

        /// <summary>
        /// Добавить пользовательский параметр рабочего окружения
        /// </summary>
        /// <param name="customItem"></param>
        public void AddCustomWorkspaceItem(WorkspaceItemBase customItem)
        {
            ItemsList.Add(customItem);
        }

        public string GetXml()
        {
            XDocument mgrXml = new XDocument();

            XElement rootElement = new XElement(
                name:nameof(WorkspaceManager),
                new XAttribute(nameof(Environment.MachineName), Environment.MachineName));
            
            foreach (WorkspaceItemBase item in WorkspaceItems)
            {
                XElement itemElement = new XElement(name:"Item",
                    new XAttribute(nameof(item.ItemType), item.ItemType.GetEnumDescription()),
                    new XElement(nameof(item.Priority), item.Priority.GetEnumDescription()),
                    new XElement(nameof(item.State), item.State.GetEnumDescription()),
                    new XElement(nameof(item.StateDetails), item.StateDetails),
                    new XElement(nameof(item.ActualValue), item.ActualValue),
                    new XElement(nameof(item.ActualIsOptimal), item.ActualIsOptimal));

                rootElement.Add(itemElement);
            }

            mgrXml.Add(rootElement);

            return mgrXml.ToString();
        }
    }
}
