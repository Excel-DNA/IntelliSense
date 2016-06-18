using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Microsoft.Win32;

namespace ExcelDna.IntelliSense.Tools
{
    public static class RegistrationFunctions
    {
        const string DisabledVersionsMachineKeyName = @"HKEY_LOCAL_MACHINE\Software\ExcelDna\IntelliSense";
        const string DisabledVersionsUserKeyName    = @"HKEY_CURRENT_USER\Software\ExcelDna\IntelliSense";
        const string DisabledVersionsValueName = "DisabledVersions";
        const string DisabledVersionsVariable = "EXCELDNA_INTELLISENSE_DISABLEDVERSIONS";

        const string ServersVariable      = "EXCELDNA_INTELLISENSE_SERVERS";
        const string ActiveServerVariable = "EXCELDNA_INTELLISENSE_ACTIVE_SERVER";

        [ExcelFunction(Description ="returns a 7x2 array of information about the IntelliSense server status")]
        public static object IntelliSenseStatus()
        {
            var machineDisabled = Registry.GetValue(DisabledVersionsMachineKeyName, DisabledVersionsValueName, null) as string;
            var userDisabled    = Registry.GetValue(DisabledVersionsUserKeyName, DisabledVersionsValueName, null) as string;
            var environmentDisabled = Environment.GetEnvironmentVariable(DisabledVersionsVariable) as string;
            var serversString       = Environment.GetEnvironmentVariable(ServersVariable) as string; // We could split on ';', then on ','
            var activeServer        = Environment.GetEnvironmentVariable(ActiveServerVariable) as string;
            var parts = activeServer?.Split(',');
            var activeRegistrationXllPath = parts?[0];
            var activeRegistrationId      = parts?[1];
            var activeRegistrationVersion = parts?[2];

            return new object[,]
                {
                    { "RegistryMachineDisabled", machineDisabled            ?? "" },
                    { "RegistryUserDisabled",    userDisabled               ?? "" },
                    { "EnvironmentDisabled",     environmentDisabled        ?? "" },
                    { "Servers",                 serversString              ?? "" },
                    { "ActiveServerXllPath",     activeRegistrationXllPath  ?? "" },
                    { "ActiveServerId",          activeRegistrationId       ?? "" },
                    { "ActiveServerVersion",     activeRegistrationVersion  ?? "" },
                };
        }
    }
}
