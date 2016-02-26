using System;
using ExcelDna.Integration;

namespace ExcelDna.IntelliSense
{
    // TODO: This is to be replaced by the Provider / Server info retrieval mechanism
    // First version might run on a timer for updates.
    public class IntelliSenseHelper : IDisposable
    {
        private readonly IntelliSenseDisplay _id;
        private readonly IIntelliSenseProvider _excelDnaProvider = new ExcelDnaIntelliSenseProvider();
        private readonly IIntelliSenseProvider _workbookProvider = new WorkbookIntelliSenseProvider();
        // TODO: Others

        public IntelliSenseHelper()
        {
            _id = new IntelliSenseDisplay();
            _id.SetXllOwner(ExcelDnaUtil.XllPath);
            RegisterIntellisense();
        }

        void RegisterIntellisense()
        {
            _excelDnaProvider.Refresh();    // Must be in macro context
            foreach (var fi in _excelDnaProvider.GetFunctionInfos())
            {
                _id.RegisterFunctionInfo(fi);
            }

            _workbookProvider.Refresh();
            foreach (var fi in _workbookProvider.GetFunctionInfos())
            {
                _id.RegisterFunctionInfo(fi);
            }
        }

        public void Dispose()
        {
            _id.Dispose();
        }
    }
}
