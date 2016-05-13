using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using ExcelDna.Integration;

namespace ExcelDna.IntelliSense
{
    // This class implements the registration and activation of this add-in as an IntelliSense Server.
    //
    // Among different add-ins that are loaded into an Excel process, at most one IntelliSenseServer can be Active.
    // This should always be the IntelliSenseServer with the greatest version number among those registered.
    // At the moment the bookkeeping for registration and activation in the process is done with environment variables. 
    // (An attractive alternative is the hidden Excel name space: http://www.cpearson.com/excel/hidden.htm )
    // This prevents cross-AppDomain calls, which are problematic because assemblies are then loaded into multiple AppDomains, and
    // since the mechanism is intended to cater for different assembly versions, this would be a problem. Also, we don't control
    // the CLR hosting configuration, so can't always set the MultiDomain flag on setup. COM mechanisms could work, but are complicated.
    // Another approach would be to use a hidden Excel function that the Active server provides, and have all server register with the active server.
    // When a new server should become active, it then tells the active server, and somehow gets all the other registrations...

    // Registered Servers also register a macro with Excel through which control calls are to be made.
    // This is against a unique GUID-based name for every registered server, so that the Activate call can be made on an inactive server.
    // (To be called in a macro context only, e.g. from AutoOpen.)

    // The Active Server also registers a macro with Excel under a well-know name (not the server-specific GUID), 
    // which passes through to a specific provider.
    // This is so that an add-in can tell (some provider in) the active server to update
    // E.g. Application.Run("ExcelDna.IntelliSense.Refresh", ' NOT WITH: "VBA") , ActiveWorkbook.Name)
    //      Application.Run("ExcelDna.IntelliSense.Refresh", ' NOT WITH: "XLL") , ExcelDnaUtil.XllPath)
    //      Application.Run("ExcelDna.IntelliSense.Refresh", ' NOT WITH: "XML") , @"C:\Temp\MyInfo.xml")
    //      XlCall.Excel(XlCall.xlcRun, "ExcelDna.IntelliSense.Refresh")
    // NB: This can't be the only way a provider knows what to load, because we don't want to do a hand-over 
    //     when a new Server becomes Active (there might not have been a server loaded at the start).
    // So we want the provider to always be able to scan to get the info.
    //     The Refresh call just allows a re-scan in a macro context (and of course would fail if no Server is Active)

    // Now:
    // When a Server becomes active it lets all the providers scan.
    // When a new add-in is then loaded, it calls the ExcelDna.IntelliSense.Refresh macro to force a re-scan.
    // When an add-in dynamically registers some more UDFs, it can call Update to force a rescan.
    // So there is no way to say "listen to this .xml file".
    // So the add-in should provide a discoverable way that a provider can call it to say "what files must I listen to?". 

    // REMEMBER: COM events are not necessarily safe macro contexts.
    public static class IntelliSenseServer
    {
        const string ServerVersion = "0.0.4";  // TODO: Define and manage this somewhere else

        // NOTE: Do not change these constants in custom versions. 
        //       They are part of the co-operative safety mechanism allowing different add-ins providing IntelliSense to work together safely.
        const string DisabledVersionsMachineKeyName = @"HKEY_LOCAL_MACHINE\Software\ExcelDna\IntelliSense";
        const string DisabledVersionsUserKeyName = @"HKEY_CURRENT_USER\Software\ExcelDna\IntelliSense";
        const string DisabledVersionsValueName = "DisabledVersions";
        const string DisabledVersionsVariable = "EXCELDNA_INTELLISENSE_DISABLEDVERSIONS";

        const string ServersVariable      = "EXCELDNA_INTELLISENSE_SERVERS";
        const string ActiveServerVariable = "EXCELDNA_INTELLISENSE_ACTIVE_SERVER";

        const string ControlMessageActivate = "ACTIVATE";
        const string ControlMessageDeactivate = "DEACTIVATE";

        // Info for registration
        // _serverId is a transient ID to identify this IntelliSense server - we could have used the ExcelDnaUtil.XllGuid one too,
        // but it wasn't public in Excel-DNA v 0.32
        // The advantage of the XllGuid one is that it would be a stable ID across runs.
        static string _xllPath = ExcelDnaUtil.XllPath;
        static Guid _serverId = Guid.NewGuid();   

        // Activation
        static bool _isActive = false;
        static IntelliSenseHelper _helper = null;

        // Called directly from AutoOpen.
        public static void Register()
        {
            TraceLogger.Initialize();

            Logger.Initialization.Info($"IntelliSenseServer.Register Begin: Version {ServerVersion} in {AppDomain.CurrentDomain.FriendlyName}");
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
            else if (RegistrationInfo.CompareVersions(ServerVersion, activeInfo.Version) > 0)
            {
                // Check version 
                // We're newer - deactivate the active server and activate ourselves.
                shouldActivate = true;
            }
            else
            {
                Logger.Initialization.Info($"IntelliSenseServer not being activated now. Active Version: {activeInfo.Version}");
            }
            // Else we're not activating - there is an active server and it is the same version or newer
            // TODO: Tell it to load our UDFs somehow - maybe call a hidden macro?

            if (shouldActivate)
            {
                var activated = Activate();
            }

            AppDomain.CurrentDomain.DomainUnload += CurrentDomain_DomainUnload;
            Logger.Initialization.Info("IntelliSenseServer.Register End");
        }

        // DomainUnload runs when AutoClose() would run on the add-in.
        // I.e when the add-in is explicitly unloaded via code or the add-ins dialog, or when the add-in is re-loaded 
        // (reload via File->Open is equivalent to unload, then load).
        // We don't expect DomainUnload to run when Excel is shutting down.
        static void CurrentDomain_DomainUnload(object sender, EventArgs e)
        {
            Logger.Initialization.Info("IntelliSenseServer DomainUnload");
            //// Early shutdown notification
            //XlCall.ShutdownStarted();

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

        // Called internally from the Register() call, or via the control function from another server.
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
            catch (Exception ex)
            {
                Logger.Initialization.Error($"IntelliSenseServer.Activate failed: {ex}");
                return false;
            }
        }

        // Called internally from the AppDomain_DomainUnload event handler, and via the control function from another server when that server figures out that it must become the active server.
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
            catch (Exception ex)
            {
                // TODO: Log
                Logger.Initialization.Error($"IntelliSenseServer.Deactivate error: {ex}");
                return false;
            }
        }

        // NOTE: Please do not remove this safety mechanism in custom versions.
        //       The IntelliSense mechanism is co-operative between independent add-ins.
        //       Allowing a safe disable options is important to support future versions, and protect against problematic bugs.
        // Checks whether this IntelliSense Server version is completely disabled
        static bool IsDisabled()
        {
            var machineDisabled = Registry.GetValue(DisabledVersionsMachineKeyName, DisabledVersionsValueName, null) as string;
            var userDisabled = Registry.GetValue(DisabledVersionsUserKeyName, DisabledVersionsValueName, null) as string;
            var environmentDisabled = Environment.GetEnvironmentVariable(DisabledVersionsVariable) as string;

            var thisVersion = ServerVersion;
            var isDisabled = IsVersionMatch(thisVersion, machineDisabled) ||
                             IsVersionMatch(thisVersion, userDisabled) ||
                             IsVersionMatch(thisVersion, environmentDisabled);

            if (isDisabled)
            {
                Logger.Initialization.Info($"IntelliSenseServer version {thisVersion} is disabled. MachineDisabled: {machineDisabled}, UserDisabled: {userDisabled}, EnvironmentDisabled: {environmentDisabled}");
            }

            return isDisabled;
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
        
        // NOTE: We have to be really careful about compatibility across versions here...
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
                catch (Exception ex)
                {
                    // TODO: Log
                    Debug.Print($"!!! ERROR: Invalid RegistrationString {registrationString}: {ex.Message}");
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
        //       Max length for an environment variable is 32,767 characters.
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
                    Version = ServerVersion 
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
            if (string.IsNullOrEmpty(versionPattern))
                return false;

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
            var servers = Environment.GetEnvironmentVariable(ServersVariable);
            if (servers == null)
            {
                Debug.Print("!!! ERROR: ServersVariable not set");
                return null;
            }
            return servers.Split(';')
                          .Select(str => RegistrationInfo.FromRegistrationString(str))
                          .Max();
        }
        #endregion

        #region IntelliSense control function registered with Excel

        static void RegisterControlFunction()
        {
            var method = typeof(IntelliSenseServer).GetMethod(nameof(IntelliSenseServerControl), BindingFlags.Static | BindingFlags.Public);
            var name = "IntelliSenseServerControl_" +_serverId.ToString("N");
            ExcelIntegration.RegisterMethods(new List<MethodInfo> { method }, 
                                             new List<object> { new ExcelCommandAttribute { Name = name } }, 
                                             new List<List<object>> { new List<object> { null } });
            // No Unregistration - that will happen automatically (and is only needed) when we are unloaded.
        }

        // NOTE: The name here is used by Reflection above (when registering the method with Excel)
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
