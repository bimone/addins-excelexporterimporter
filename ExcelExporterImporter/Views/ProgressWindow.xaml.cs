#region Using Directives

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using ExcelExporterImporter.ViewModels;
using ExcelExporterImporter.Views;

#endregion

namespace BIMOneAddinManager.Views
{
    /// <summary>
    ///     Interaction logic for Downloader.xaml
    /// </summary>
    public partial class ProgressWindow : Window
    {
        private const int GWL_STYLE = -16;
        private const int WS_SYSMENU = 0x80000;
        private readonly ProgressViewModel progressViewModel;
        private readonly ButtonUserControl userControl;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private void Window_Loaded_1(object sender, RoutedEventArgs e)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) & ~WS_SYSMENU);
        }

        #region Constructor/Destructor

        public ProgressWindow(ProgressViewModel vm)
        {
            DataContext = vm;
            progressViewModel = vm;
            InitializeComponent();

            userControl = new ButtonUserControl();
            userControl.Command = vm.Command;
            vm.PropertyChanged += vm_PropertyChanged;

            WinformHost.Child = userControl;
        }

        private void vm_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case "ButtonEnabled":
                    userControl.Enabled = progressViewModel.ButtonEnabled;
                    break;
                case "ButtonText":
                    userControl.ButtonText = progressViewModel.ButtonText;
                    break;
            }
        }

        #endregion
    }
}