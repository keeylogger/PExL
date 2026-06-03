using ExcelDna.Integration;

namespace PExL.AddIn
{
    /// <summary>
    /// Entry point recognized by Excel-DNA. Wires up selection tracking on load.
    /// </summary>
    public sealed class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Interop.SelectionTracker.Install();
        }

        public void AutoClose()
        {
            Interop.SelectionTracker.Uninstall();
        }
    }
}
