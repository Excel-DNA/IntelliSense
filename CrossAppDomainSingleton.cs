using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDna.IntelliSense
{
    class CrossAppDomainSingleton : MarshalByRefObject
    {
        private const string AppDomainName = "ExcelDNASingletonAppDomain";

        private static IntelliSenseDisplay _intelliSenseDisplay;
      
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

        public static IntelliSenseDisplay GetOrCreate()
        {
            Type type = typeof(IntelliSenseDisplay);

            AppDomain appDomain = GetAppDomain(AppDomainName);

            var xllName = Win32Helper.GetXllName();

            if (appDomain == null)
            {
                AppDomainSetup domaininfo = new AppDomainSetup();
                domaininfo.ApplicationBase = Path.GetDirectoryName(xllName);
                appDomain = AppDomain.CreateDomain(AppDomainName, AppDomain.CurrentDomain.Evidence, domaininfo);
            }

            _intelliSenseDisplay = (IntelliSenseDisplay)appDomain.GetData(type.FullName);

            if (_intelliSenseDisplay == null)
            {
                _intelliSenseDisplay = (IntelliSenseDisplay)appDomain.CreateInstanceAndUnwrap(type.Assembly.FullName, type.FullName);
                _intelliSenseDisplay.SetXllOwner(Win32Helper.GetXllName());
            }

            _intelliSenseDisplay.AddReference(xllName);

            return _intelliSenseDisplay;
        }

        public static void RemoveReference()
        {
            _intelliSenseDisplay.RemoveReference(Win32Helper.GetXllName());

            if (!_intelliSenseDisplay.IsUsed())
            {
                _intelliSenseDisplay.Shutdown();
                _intelliSenseDisplay = null;
            }
        }
    }
}
