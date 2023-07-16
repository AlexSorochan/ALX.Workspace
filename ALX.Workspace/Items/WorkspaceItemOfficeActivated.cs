using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemOfficeActivated : WorkspaceItemBase
    {
        private const string
            ActivatedStr = "Активирован",
            NotActivatedStr = "Не активирован",
            HelpLink = "https://docs.microsoft.com/ru-ru/office365/troubleshoot/licensing/determine-office-license-type?tabs=windows",
            VbsFileName = "ospp.vbs",
            CmdCommand = "cscript",
            VbsArgument = "/dstatus";

        private string _xlLicResult;

        public override WorkspaceItem ItemType => WorkspaceItem.OfficeActivation;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.High;
        public override string Help => 
            $"Проверка лицензии Office, по <href={HelpLink}>рекомендациям Microsoft</href>:{Environment.NewLine}" +
            $"К примеру, в случае с MS Office 2013{Environment.NewLine}" +
            $"[c:\\Program Files\\Microsoft Office\\Office15\\]{Environment.NewLine}" +
            $"CMD: [{CmdCommand} {VbsFileName} {VbsArgument}]{Environment.NewLine}" +
            $"{Environment.NewLine}OFFICE LICENSE:{Environment.NewLine}{_xlLicResult}";

        public override string[] OptimalValues => new[] { ActivatedStr, NotActivatedStr };

        public override bool ActualIsOptimal => string.Equals(ActualValue, ActivatedStr);

        public override bool UpdateStateAction(out string actualValue)
        {
            string xlPath;
            actualValue = string.Empty;

            #region Получить Excel.Path

            Application application = null;

            try
            {
                application = new Application();
                xlPath = application.Path;
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

            // excel.exe располагается в \root\office<number>\
            // в случае с некоторыми версиями офиса
            // убрать "\root" из пути
            string rootPathPart = @"\root";
            if (xlPath.ToLower().Contains(rootPathPart)) 
                xlPath = xlPath.ToLower().Replace(rootPathPart, string.Empty);

            string vbsPath = Path.Combine(xlPath, VbsFileName);
            if (!File.Exists(vbsPath))
            {
                _xlLicResult = $"Не удалось найти файл: [{vbsPath}]";
                return false;
            }

            #endregion

            #region Пример результат выполнения команды:

            /* Сервер сценариев Windows (Microsoft R) версия 5.812
             * Copyright (C) Корпорация Майкрософт 1996-2006, все права защищены.
             * 
             * ---Processing--------------------------
             * ---------------------------------------
             * SKU ID: b322da9c-a2e2-4058-9e4e-f59a6970bd69
             * LICENSE NAME: Office 15, OfficeProPlusVL_KMS_Client edition
             * LICENSE DESCRIPTION: Office 15, VOLUME_KMSCLIENT channel
             * LICENSE STATUS:  ---LICENSED--- 
             * ERROR CODE: 0x4004F040 (for information purposes only as the status is licensed)
             * ERROR DESCRIPTION: The Software Licensing Service reported that the product was activated but the owner should verify the Product Use Rights.
             * REMAINING GRACE: 178 days  (256748 minute(s) before expiring)
             * Last 5 characters of installed product key: GVGXT
             * Activation Type Configuration: ALL
	         * KMS machine name from DNS: ftp.agroprombank.com:1688
	         * Activation Interval: 120 minutes
	         * Renewal Interval: 10080 minutes
	         * KMS host caching: Enabled
             * ---------------------------------------
             * ---------------------------------------
             * ---Exiting-----------------------------
             */

            #endregion

            #region Проверка лицензии Office CMD:[cscript ospp.vbs /dstatus]

            int timeout = 15000;
            StringBuilder outputBuilder = new StringBuilder();

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                // 866 - Russian
                StandardOutputEncoding = Encoding.GetEncoding(866),
                WorkingDirectory = xlPath,
                Arguments = $"{VbsFileName} {VbsArgument}",
                FileName = CmdCommand,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                // для того чтобы скрыть консольное окно
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardError = true,

            };
            
            using (AutoResetEvent outputWaitHandle = new AutoResetEvent(false))
            {
                using (Process process = new Process
                {
                    StartInfo = startInfo,
                    // для того чтобы скрыть консольное окно
                    EnableRaisingEvents = true
                })
                {
                    try
                    {
                        process.OutputDataReceived += (sender, e) =>
                        {
                            if (e.Data == null)
                            {
                                // ReSharper disable once AccessToDisposedClosure
                                outputWaitHandle.Set();
                            }
                            else
                            {
                                outputBuilder.AppendLine(e.Data);
                            }
                        };

                        process.Start();
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();
                        process.WaitForExit(timeout);

                        _xlLicResult = outputBuilder.ToString();
                    }
                    finally
                    {
                        outputWaitHandle.WaitOne(timeout);
                    }
                }
            }

            #endregion

            string licensedPattern = @"LICENSE\s+STATUS\:\s+[\-]{1,}LICENSED[\-]{1,}";
            bool isLicensed = Regex.IsMatch(_xlLicResult, licensedPattern);
            actualValue = isLicensed ? ActivatedStr : NotActivatedStr;
            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
