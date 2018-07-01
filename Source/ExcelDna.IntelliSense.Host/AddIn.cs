using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace TestAddIn
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
            //ExcelDna.Logging.LogDisplay.Show();
            //ExcelDna.Logging.LogDisplay.DisplayOrder = ExcelDna.Logging.DisplayOrder.NewestFirst;
        }

        public void AutoClose()
        {
            // The explicit Uninstall call was added in version 1.0.10, 
            // to give us a chance to unhook the WinEvents (which needs to happen on the main thread)
            // No easy plan to do this from the AppDomain unload event, which runs on a ThreadPool thread.
            IntelliSenseServer.Uninstall();
        }
    }
}
