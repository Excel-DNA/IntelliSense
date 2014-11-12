using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.IntelliSense;

namespace TestAddIn
{
    public static class TestFunctions
    {
        [ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
        public static double AddThem(
            [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] double v1,
            [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     double v2)
        {
            return v1 + v2;
        }

        [Description("Test function for the amazing Excel-DNA IntelliSense feature")]
        public static string jDummyFunc()
        {
            return "Howzit !";
        }

    }

    public class AddIn : IExcelAddIn
    {
        //ShutdownHelper _shutdownHelper;
        IntelliSenseDisplay _id;
        public void AutoOpen()
        {
            _id = CrossAppDomainSingleton.GetOrCreate();
            
            //_shutdownHelper = new ShutdownHelper();
            //ExcelComAddInHelper.LoadComAddIn(_shutdownHelper);

            RegisterIntellisense();
        }

        public void AutoClose()
        {
            CrossAppDomainSingleton.RemoveReference();
        }

        void RegisterIntellisense()
        {
            // Get function info from Excel-DNA, and register 
            var regInfos = ExcelIntegration.GetFunctionRegistrationInfo();
            foreach (var regInfo in regInfos)
            {
                _id.RegisterFunctionInfo(GetFunctionInfo(regInfo));
            }

            // Explicit function info registration (imagine getting from an xml file or something)
            /*
            _id.RegisterFunctionInfo(new IntelliSenseFunctionInfo 
                { 
                    FunctionName = "AddThem", 
                    Description = "A useful test function that adds two numbers, and returns the sum.",
                    ArgumentList = new List<IntelliSenseFunctionInfo.ArgumentInfo>
                    {
                        new IntelliSenseFunctionInfo.ArgumentInfo { ArgumentName = "augend", Description = "is the first value, to which will be added" },
                        new IntelliSenseFunctionInfo.ArgumentInfo { ArgumentName = "addend", Description = "is the second that will be added" },
                    }
                });
            _id.RegisterFunctionInfo(new IntelliSenseFunctionInfo 
                { 
                    FunctionName = "jDummy", 
                    Description = "Test function for the amazing Excel-DNA IntelliSense feature!" ,
                    ArgumentList = new List<IntelliSenseFunctionInfo.ArgumentInfo>()
                });
            */ 
        }

        IntelliSenseFunctionInfo GetFunctionInfo(List<string> regInfo)
        {
            var xllName = Win32Helper.GetXllName();
            
            // name, category, helpTopic, argumentNames, [description], [argumentDescription_1] ... [argumentDescription_n]
            var funcName = regInfo[0];
            var funcDesc = regInfo.Count >= 5 ? regInfo[4] : (funcName + " function");
            var argNames = regInfo[3].Split(',');
            var funcInfo = new IntelliSenseFunctionInfo
            {
                FunctionName = funcName,
                Description = funcDesc,
                ArgumentList = new List<IntelliSenseFunctionInfo.ArgumentInfo>
                    (argNames.Select((argName, i) => new IntelliSenseFunctionInfo.ArgumentInfo
                        {
                            ArgumentName = argName,
                            Description = regInfo.Count >= 6 + i ? regInfo[5+i] : ""
                        })),
                XllPath = xllName
            };
            return funcInfo;
        }
    }

    //[ComVisible(true)]
    //public class ShutdownHelper : ExcelComAddIn
    //{
    //    public override void OnConnection(object Application, ExcelDna.Integration.Extensibility.ext_ConnectMode ConnectMode, object AddInInst, ref System.Array custom)
    //    {
    //        Debug.Print("ShutdownHelper loaded.");
    //    }

    //    public override void OnDisconnection(ExcelDna.Integration.Extensibility.ext_DisconnectMode RemoveMode, ref System.Array custom)
    //    {
            IntelliSenseDisplay.Shutdown();
    //    }
    //}

}
