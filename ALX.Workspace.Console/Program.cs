using System;
using System.ComponentModel;

namespace ALX.Workspace.Console
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                System.Console.Clear();
                System.Console.WriteLine($"Проверка параметров рабочего окружения ...:{Environment.NewLine}");

                WorkspaceManager manager = WorkspaceManager.Initialize(
                    checkWindowsVersion: true, checkFrameworkVersion: true, checkFramework472: true,
                    checkFramework48: true, checkVsto2010: true, checkOfficeVersion: true,
                    checkOfficeActivated: true, checkCultureRu: true, checkWindowsSeparator: true,
                    checkOfficeSeparator: true, checkRootCerts: true);

                // Добавляем кастомную проверку
                //CustomWorkspaceItem customWorkspaceItem = new CustomWorkspaceItem();
                //manager.AddCustomWorkspaceItem(customWorkspaceItem);

                manager.Update();

                System.Console.Clear();
                System.Console.WriteLine($"Параметры рабочего окружения:{Environment.NewLine}");

                for (int i = 0; i < manager.WorkspaceItems.Length; i++)
                    System.Console.WriteLine($"{i + 1}) {manager.WorkspaceItems[i]}");

                System.Console.ReadKey();
            }
            catch (Exception exception)
            {
                if (exception is WarningException warning) System.Console.WriteLine($"WARNING => {warning.Message}");
                else System.Console.WriteLine($"ERROR => {exception}");
            }

        }
    }
}
