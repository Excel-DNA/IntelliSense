using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using ExcelDna.Integration;

namespace ExcelDna.IntelliSense
{
    // This class implements the registration and activation of this add-in as an IntelliSense Server.
    public static class IntelliSenseServer
    {
        const string _version = "0.0.1";  // TODO: Define and manage this somewhere else

        const string ServersVariable      = "EXCELDNA_INTELLISENSE_SERVERS";
        const string ActiveServerVariable = "EXCELDNA_INTELLISENSE_ACTIVE_SERVER";

        const string ControlMessageActivate = "ACTIVATE";
        const string ControlMessageDeactivate = "DEACTIVATE";

        // A transient ID to identify this IntelliSense server - we could have used the ExcelDnaUtil.XllGuid one too,
        // but it wasn't public in Excel-DNA v 0.32
        // Info for registration
        static string _xllPath = ExcelDnaUtil.XllPath;
        static Guid _serverId = Guid.NewGuid();   

        // Activation
        static bool _isActive = false;
        static IntelliSenseHelper _helper = null;

        // Called directly (not via reflection) from AutoOpen.
        public static void Register()
        {
            if (IsDisabled())
                return;

            RegisterControlFunction();
            PublishRegistration();

            bool shouldActivate = false;
            var activeInfo = GetActiveRegistrationInfo();
            if (activeInfo == null)
            {
                shouldActivate = true;
            }
            else if (RegistrationInfo.CompareVersions(_version, activeInfo.Version) > 0)
            {
                // Check version 
                // We're newer - deactivate the active server and activate ourselves.
                shouldActivate = true;
            }
            // Else we're not activating - there is an active server and it is the same version or newer
            // TODO: Tell it to load our UDFs somehow - maybe call a hidden macro?

            if (shouldActivate)
            {
                var activated = Activate();
            }

            AppDomain.CurrentDomain.DomainUnload += CurrentDomain_DomainUnload;
        }

        static void CurrentDomain_DomainUnload(object sender, EventArgs e)
        {
            UnpublishRegistration();
            if (_isActive)
            {
                Deactivate();

                var highestRegistration = GetHighestPublishedRegistration();
                if (highestRegistration != null)
                {
                    ActivateServer(highestRegistration);
                }
            }
        }

        // Called internally from the Register() call, or via reflection from another server.
        internal static bool Activate()
        {
            try
            {
                SetActiveRegistrationInfo();
                _isActive = true;

                // Now initialize (TODO: perhaps lazily...?)
                _helper = new IntelliSenseHelper();
                // TODO: Perhaps also register macro to trigger updates
                return true;
            }
            catch (Exception /*ex*/)
            {
                // TOOD: Log
                return false;
            }
        }

        // Called internally from the AppDomain_DomainUnload event handler, and via reflection from another server when that server figures out that it must become the active server.
        internal static bool Deactivate()
        {
            try
            {
                if (_helper != null)
                    _helper.Dispose();

                _isActive = false;
                ClearActiveRegistrationInfo();
                return true;
            }
            catch (Exception /*ex*/)
            {
                // TODO: Log
                return false;
            }
        }

        static bool IsDisabled()
        {
            // TODO: Check where this version is disabled, by checking the reigstry and environement variables
            // var machineDisabled = Registry.GetValue()
            // var userDisabled = Registry.GetValue()
            // var environmentDisabled = Environment.GetEnvironmentVariable(...)
            return false;
        }

        // Attempts to activate the server described by registrationInfo
        // return true for success, false for any problems
        static bool ActivateServer(RegistrationInfo registrationInfo)
        {
            // Suppress errors if things go wrong, including unexpected return types.
            try
            {
                var result = ExcelDna.Integration.XlCall.Excel(ExcelDna.Integration.XlCall.xlfCall, registrationInfo.GetControlMacroName(), ControlMessageDeactivate);
                return (bool)result;
            }
            catch (Exception /*ex*/)
            {
                // TODO: Log
                return false;
            }
        }

        // Attempts to deactivate the server described by registrationInfo
        // returns true for success, false if there were problems, 
        static bool DeactivateServer(RegistrationInfo registrationInfo)
        {
            // Suppress errors if things go wrong, including unexpected return types.
            try
            {
                var result = ExcelDna.Integration.XlCall.Excel(ExcelDna.Integration.XlCall.xlfCall, registrationInfo.GetControlMacroName(), ControlMessageDeactivate);
                return (bool)result;
            }
            catch (Exception /*ex*/)
            {
                // TODO: Log
                return false;
            }
        }

        #region Registration
        
        // NOTE: We have to be really careful about compatibility here...
        class RegistrationInfo : IComparable<RegistrationInfo>
        {
            public string XllPath;
            public Guid ServerId;
            public string Version;

            public static RegistrationInfo FromRegistrationString(string registrationString)
            {
                try
                {
                    var parts = registrationString.Split(',');
                    return new RegistrationInfo
                    {
                        XllPath = parts[0],
                        ServerId = Guid.ParseExact(parts[1], "N"),
                        Version = parts[2]
                    };
                }
                catch
                {
                    // TODO: Log
                    return null;
                }
            }

            public string ToRegistrationString()
            {
                return string.Join(",", XllPath, ServerId.ToString("N"), Version);
            }

            public int CompareTo(RegistrationInfo other)
            {
                return CompareVersions(Version, other.Version);
            }

            public string GetControlMacroName()
            {
                return "IntelliSenseControl_" + ServerId.ToString("N");
            }

            // 1.2.0 is equal to 1.2
            // Returns: -1 if version1 < version2
            //          0 if version1 == version2
            //          1 if version1 > version2
            // Invalid version strings are considered very small (as if they are 0
            public static int CompareVersions(string versionString1, string versionString2)
            {
                int[] version1 = ParseVersion(versionString1);
                int[] version2 = ParseVersion(versionString2);

                var maxLength = Math.Max(version1.Length, version2.Length);
                for (int i = 0; i < maxLength; i++)
                {
                    int v1 = (version1.Length - 1) < i ? 0 : version1[i];
                    int v2 = (version2.Length - 1) < i ? 0 : version2[i];
                    if (v1 < v2)
                        return -1;
                    if (v1 > v2)
                        return 1;
                }
                return 0;
            }
        }

        // NOTE: We assume this will always run on the main thread in the process, so we have no synchronization.
        //       Max length for en environment variable is 32,767 characters.
        static void PublishRegistration()
        {
            var oldServers = Environment.GetEnvironmentVariable(ServersVariable);
            var newServers = (oldServers == null) ? RegistrationString : string.Join(";", oldServers, RegistrationString);
            Environment.SetEnvironmentVariable(ServersVariable, newServers);
        }

        static void UnpublishRegistration()
        {
            var oldServers = new List<string>(Environment.GetEnvironmentVariable(ServersVariable).Split(';'));
            var removed = oldServers.Remove(RegistrationString);
            Debug.Assert(removed, ("IntelliSenseServer.UnpublishRegistration - Registration not found in " + ServersVariable));
            var newServers = string.Join(";", oldServers);
            Environment.SetEnvironmentVariable(ServersVariable, newServers);
        }

        // returns null if there is no active IntelliSense Server.
        static RegistrationInfo GetActiveRegistrationInfo()
        {
            var activeString = Environment.GetEnvironmentVariable(ActiveServerVariable);
            if (string.IsNullOrEmpty(activeString))
                return null;
            return RegistrationInfo.FromRegistrationString(activeString);
        }

        static void SetActiveRegistrationInfo()
        {
            var currentActive = Environment.GetEnvironmentVariable(ActiveServerVariable);
            Debug.Assert(currentActive == null, "ActiveServer already set while activating");
            Environment.SetEnvironmentVariable(ActiveServerVariable, RegistrationString);
        }

        static void ClearActiveRegistrationInfo()
        {
            Environment.SetEnvironmentVariable(ActiveServerVariable, null);
        }

        static string RegistrationString
        {
            get
            {
                var ri = new RegistrationInfo 
                { 
                    XllPath = ExcelDnaUtil.XllPath,
                    ServerId = _serverId,
                    Version = _version 
                };
                return ri.ToRegistrationString();
            }
        }

        // Versions are dotted integer strings, e.g. 1.2.3
        // Invalid strings parse as [0]
        static int[] ParseVersion(string versionString)
        {
            if (string.IsNullOrEmpty(versionString))
            {
                return new int[] {0};
            }
            var versionParts = versionString.Split('.');
            int[] version = new int[versionParts.Length];
            for (int i = 0; i < versionParts.Length; i++)
            {
                int versionPart;
                if (!int.TryParse(versionParts[i], out versionPart))
                {
                    return new int[] {0};
                }
                version[i] = versionPart;
            }
            return version;
        }

        // Version patterns are ","-joined lists of (dotted integer strings, with a possible trailing .* wildcard).
        // e.g. 1.2.*, which would be matched with regex 1\.2(\.\d+)*
        static bool IsVersionMatch(string version, string versionPattern)
        {
            var regexParts = new List<string>();
            var parts = versionPattern.Split(',');
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (Regex.IsMatch(trimmed, @"(\d+\.)+(\.\*)?", RegexOptions.None))
                {
                    regexParts.Add("^" + trimmed.Replace(".*", @"(\.\d+)*") + "$");
                }
            }
            var regex = string.Join(@"|", regexParts);
            return Regex.IsMatch(version, regex);
        }

        // returns null if there are none registered
        static RegistrationInfo GetHighestPublishedRegistration()
        {
            return Environment.GetEnvironmentVariable(ServersVariable)
                              .Split(';')
                              .Select(str => RegistrationInfo.FromRegistrationString(str))
                              .Max();
        }
        #endregion

        //#region AppDomain helper
        //static AppDomain GetAppDomain(string friendlyName)
        //{
        //    IntPtr enumHandle = IntPtr.Zero;
        //    mscoree.ICorRuntimeHost host = new mscoree.CorRuntimeHost();

        //    try
        //    {
        //        host.EnumDomains(out enumHandle);

        //        while (true)
        //        {
        //            object domain;
        //            host.NextDomain(enumHandle, out domain);

        //            if (domain == null)
        //                break;

        //            AppDomain appDomain = (AppDomain)domain;
        //            if (appDomain.FriendlyName.Equals(friendlyName))
        //                return appDomain;
        //        }
        //    }
        //    finally
        //    {
        //        host.CloseEnum(enumHandle);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(host); // Why??? Pure cargo-culting here...
        //    }

        //    return null;
        //}
        //#endregion

        #region IntelliSense control function registered with Excel

        static void RegisterControlFunction()
        {
            var method = typeof(IntelliSenseServer).GetMethod("IntelliSenseServerControl", BindingFlags.Static | BindingFlags.Public);
            var name = "IntelliSenseServerControl_" +_serverId.ToString("N");
            ExcelIntegration.RegisterMethods(new List<MethodInfo> { method }, 
                                             new List<object> { new ExcelCommandAttribute { Name = name } }, 
                                             new List<List<object>> { new List<object> { null } });
            // No Unregistration - that will happen when we are unloaded.
        }

        public static object IntelliSenseServerControl(object control)
        {
            if (control is string && (string)control == ControlMessageActivate)
            {
                Debug.Print("IntelliSenseServer.Activate in AppDomain: " + AppDomain.CurrentDomain.FriendlyName);
                return IntelliSenseServer.Activate();
            }
            else if (control is string && (string)control == ControlMessageDeactivate)
            {
                Debug.Print("IntelliSenseServer.Deactivate in AppDomain: " + AppDomain.CurrentDomain.FriendlyName);
                return IntelliSenseServer.Deactivate();
            }
            return false;
        }

        #endregion
    }
}
