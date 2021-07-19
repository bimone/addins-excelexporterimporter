using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Windows;
using Autodesk.Revit.DB;
using ExcelExporterImporter.ViewModels;

namespace ExcelExporterImporter.Views
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly ConsoleTraceListener consoleTraceListener;

        public MainWindow(Document doc)
        {
            //====================================On applique la même langue que Revit/We use the same language as Revit=========================================
            var application = doc.Application;
            var lang = application.Language;
            if (lang.ToString().Contains("French"))
            {
                var cultureInfo = new CultureInfo("fr-FR");
                Thread.CurrentThread.CurrentCulture = Thread.CurrentThread.CurrentUICulture = cultureInfo;
            }

            //====================================================================================================================================================
            InitializeComponent();

            var mainViewModel = new MainViewModel(this, doc);
            // this creates an instance of the ViewModel
            DataContext = mainViewModel; // this sets the newly created ViewModel as the DataContext for the View
            if (mainViewModel.CloseAction == null)
                mainViewModel.CloseAction = Close;

            consoleTraceListener = new ConsoleTraceListener();
            //   this.debugTraceListener = new DebugTraceListener();

            PresentationTraceSources.Refresh();
            PresentationTraceSources.DataBindingSource.Listeners.Add(consoleTraceListener);
            // PresentationTraceSources.DataBindingSource.Listeners.Add(debugTraceListener);
            PresentationTraceSources.DataBindingSource.Switch.Level = SourceLevels.Warning | SourceLevels.Error;
        }
        //  private readonly DebugTraceListener debugTraceListener;

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            if (!e.Cancel) Settings.Default.Save();
        }

        protected override void OnClosed(EventArgs e)
        {
            PresentationTraceSources.DataBindingSource.Listeners.Remove(consoleTraceListener);
            //  PresentationTraceSources.DataBindingSource.Listeners.Remove(this.debugTraceListener);
            PresentationTraceSources.Refresh();
            base.OnClosed(e);
        }

    }
}