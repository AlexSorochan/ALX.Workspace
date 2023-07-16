using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

namespace ALX.Workspace.Items
{
    public class WorkspaceItemRootCerts : WorkspaceItemBase
    {
        private const string
            RootPmrCaName = "ROOT-PMR-CA",
            RootPmrCaThumbprint = "63D01C855A2EFA83CC85E6A9E062AE6595A958E3";

        public override WorkspaceItem ItemType => WorkspaceItem.RootCerts;
        public override string DisplayName => ItemType.GetEnumDescription();
        public override WorkspaceItemPriority Priority => WorkspaceItemPriority.High;
        
        public override string Help => 
            $"Проверка наличия корневых сертификатов в хранилище пользователя:{Environment.NewLine}" +
            $"1) [Win] + [R]{Environment.NewLine}" +
            $"2) Выполнить команду: [certmgr.msc]{Environment.NewLine}" +
            $"3) Перейти в хранилище: [Доверенные корневые центры сертификации] >> [Сертификаты]{Environment.NewLine}" +
            $"4) Убедиться в наличие сертификата: \"{RootPmrCaName}\"";

        public override string[] OptimalValues => new[] { RootPmrCaName };

        public override bool ActualIsOptimal => OptimalValues.All(x => ActualValue.Contains(x));

        public override bool UpdateStateAction(out string actualValue)
        {
            actualValue = string.Empty;

            #region ROOT-PMR-CA

            using (X509Store store = new X509Store(StoreName.Root, StoreLocation.LocalMachine))
            {
                store.Open(OpenFlags.ReadOnly);

                X509CertificateCollection collection = store.Certificates.
                    Find(X509FindType.FindByTimeValid, DateTime.Now, true).
                    Find(X509FindType.FindByThumbprint, RootPmrCaThumbprint, true);

                X509Certificate2 certificate = collection.Cast<X509Certificate2>().FirstOrDefault();

                if (certificate != null)
                    actualValue += $"{RootPmrCaName}; ";

                store.Close();
            }

            #endregion

            actualValue = actualValue.TrimEnd();
            return !string.IsNullOrEmpty(actualValue);
        }
    }
}
