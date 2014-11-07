using System;
using System.Runtime.InteropServices;

namespace ExcelDna.IntelliSense
{
    class CrossAppDomainSingleton : MarshalByRefObject
    {
        private const string AppDomainName = "ExcelDNASingletonAppDomain";

        private IntelliSenseDisplay intelliSenseDisplay;
        
        // Note : it might be useful in the future to know what xll created the IntelliSenseDisplay instance
        // Not used yet.
        private string intelliSenseDisplayDomain;

        public IntelliSenseDisplay IntelliSenseDisplay
        {
            get { return intelliSenseDisplay; }
        }

        public void SetIntelliSenseDisplay(IntelliSenseDisplay intelliSense, string appDomainName)
        {
            intelliSenseDisplay = intelliSense;
            intelliSenseDisplayDomain = appDomainName;
        }

        private static AppDomain GetAppDomain(string friendlyName)
        {
            IntPtr enumHandle = IntPtr.Zero;
            mscoree.CorRuntimeHost host = new mscoree.CorRuntimeHost();

            try
            {
                host.EnumDomains(out enumHandle);

                while (true)
                {
                    object domain;
                    host.NextDomain(enumHandle, out domain);

                    if (domain == null)
                    {
                        break;
                    }

                    AppDomain appDomain = (AppDomain) domain;
                    if (appDomain.FriendlyName.Equals(friendlyName))
                    {
                        return appDomain;
                    }
                }
            }
            finally
            {
                host.CloseEnum(enumHandle);
                Marshal.ReleaseComObject(host);
            }

            return null;
        }

        // Note : the IntelliSenseDisplay instance will belong to the first xll
        // if that xll is unloaded, then the instance will be disposed, but this piece
        // of code will still return that instance
        public static IntelliSenseDisplay GetOrCreate()
        {
            Type type = typeof(CrossAppDomainSingleton);

            AppDomain appDomain = GetAppDomain(AppDomainName);

            if (appDomain == null)
            {
                AppDomainSetup domaininfo = new AppDomainSetup();
                domaininfo.ApplicationBase = Environment.CurrentDirectory;

                appDomain = AppDomain.CreateDomain(AppDomainName, AppDomain.CurrentDomain.Evidence, domaininfo);
            }

            CrossAppDomainSingleton instance = (CrossAppDomainSingleton)appDomain.GetData(type.FullName);

            if (instance == null)
            {
                instance = (CrossAppDomainSingleton)appDomain.CreateInstanceAndUnwrap(type.Assembly.FullName, type.FullName);
                appDomain.SetData(type.FullName, instance);
            }

            IntelliSenseDisplay intelliSense = instance.IntelliSenseDisplay;

            if (intelliSense == null)
            {
                intelliSense = new IntelliSenseDisplay();
                instance.SetIntelliSenseDisplay(intelliSense, AppDomain.CurrentDomain.FriendlyName);
            }

            return intelliSense;
        }
    }
}
