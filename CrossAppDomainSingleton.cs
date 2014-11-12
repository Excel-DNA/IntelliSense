using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDna.IntelliSense
{
    class CrossAppDomainSingleton : MarshalByRefObject
    {
        private const string AppDomainName = "ExcelDNASingletonAppDomain";

        private static CrossAppDomainSingleton _instance;

        private IntelliSenseDisplay _intelliSenseDisplay;
        
        // Note : it might be useful in the future to know what xll created the IntelliSenseDisplay instance
        // Not used yet.
        private string _intelliSenseDisplayDomain;

        public IntelliSenseDisplay IntelliSenseDisplay
        {
            get { return _intelliSenseDisplay; }
        }

        public void SetIntelliSenseDisplay(IntelliSenseDisplay intelliSense, string appDomainName)
        {
            _intelliSenseDisplay = intelliSense;
            _intelliSenseDisplayDomain = appDomainName;
        }

        public void Reset()
        {
            IntelliSenseDisplay.Shutdown();
            _intelliSenseDisplay = null;
            _intelliSenseDisplayDomain = null;
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

            var xllName = Win32Helper.GetXllName();

            if (appDomain == null)
            {
                AppDomainSetup domaininfo = new AppDomainSetup();
                domaininfo.ApplicationBase = Path.GetDirectoryName(xllName);
                appDomain = AppDomain.CreateDomain(AppDomainName, AppDomain.CurrentDomain.Evidence, domaininfo);
            }

            _instance = (CrossAppDomainSingleton)appDomain.GetData(type.FullName);

            if (_instance == null)
            {
                _instance = (CrossAppDomainSingleton)appDomain.CreateInstanceAndUnwrap(type.Assembly.FullName, type.FullName);
                appDomain.SetData(type.FullName, _instance);
            }

            IntelliSenseDisplay intelliSense = _instance.IntelliSenseDisplay;

            if (intelliSense == null)
            {
                intelliSense = new IntelliSenseDisplay();
                _instance.SetIntelliSenseDisplay(intelliSense, AppDomain.CurrentDomain.FriendlyName);
            }

            intelliSense.AddReference(xllName);

            return intelliSense;
        }

        public static void RemoveReference()
        {
            _instance.IntelliSenseDisplay.RemoveReference(Win32Helper.GetXllName());

            if (!_instance.IntelliSenseDisplay.IsUsed())
            {
                _instance.Reset();
            }
        }
    }
}
