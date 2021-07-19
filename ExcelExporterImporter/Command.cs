using System;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using ExcelExporterImporter.Views;
using log4net;
using IWin32Window = System.Windows.Forms.IWin32Window;

namespace ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        private static readonly ILog Logger =
            LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public Result Execute(ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                if (doc == null)
                    return Result.Cancelled;
                var dlg = new MainWindow(doc);

                var window = new WindowInteropHelper(dlg);
                window.Owner = ComponentManager.ApplicationWindow;
                dlg.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                Logger.Error(e.Message);
                return Result.Failed;
            }
        }
    }

    internal class WindowWrapper : IWin32Window, IDisposable
    {
        public WindowWrapper(IntPtr handle)
        {
            Handle = handle;
        }

        public void Dispose()
        {
            Handle = IntPtr.Zero;
        }

        public IntPtr Handle { get; private set; }
    }
}