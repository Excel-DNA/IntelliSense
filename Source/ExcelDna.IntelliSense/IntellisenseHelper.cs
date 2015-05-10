using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace ExcelDna.IntelliSense
{
    // TODO: This is to be replaced by the Provider / Server info retrieval mechanism
    public class IntelliSenseHelper : IDisposable
    {
        private readonly IntelliSenseDisplay _id;

        public IntelliSenseHelper()
        {
            _id = new IntelliSenseDisplay();
            _id.SetXllOwner(ExcelDnaUtil.XllPath);
            RegisterIntellisense();
        }

        void RegisterIntellisense()
        {
            string xllPath = Win32Helper.GetXllName();
            object[,] regInfos = ExcelIntegration.GetRegistrationInfo(xllPath, -1) as object[,];

            if (regInfos != null)
            {
                for (int i = 0; i < regInfos.GetLength(0); i++)
                {
                    if (regInfos[i, 0] is ExcelDna.Integration.ExcelEmpty)
                    {
                        string functionName = regInfos[i, 3] as string;
                        string description = regInfos[i, 9] as string;

                        string argumentStr = regInfos[i, 4] as string;
                        string[] argumentNames = string.IsNullOrEmpty(argumentStr) ? new string[0] : argumentStr.Split(',');

                        List<IntelliSenseFunctionInfo.ArgumentInfo> argumentInfos = new List<IntelliSenseFunctionInfo.ArgumentInfo>();
                        for (int j = 0; j < argumentNames.Length; j++)
                        {
                            argumentInfos.Add(new IntelliSenseFunctionInfo.ArgumentInfo { ArgumentName = argumentNames[j], Description = regInfos[i, j + 10] as string });
                        }

                        _id.RegisterFunctionInfo(new IntelliSenseFunctionInfo
                        {
                            FunctionName = functionName,
                            Description = description,
                            ArgumentList = argumentInfos,
                            XllPath = xllPath
                        });
                    }
                }
            }
        }

        public void Dispose()
        {
            _id.Dispose();
        }
    }
}
