using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using IWin32Window = System.Windows.Forms.IWin32Window;

namespace ExcelExporterImporter.Common
{
    public static class CommonDialogExtensions
    {
        /// <summary>
        ///     Displays a dialogbox
        /// </summary>
        /// <param name="dialog"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        public static DialogResult ShowDialog(this CommonDialog dialog, Window parent)
        {
            return dialog.ShowDialog(new Wpf32Window(parent));
        }

        public class Wpf32Window : IWin32Window
        {
            public Wpf32Window(Window wpfWindow)
            {
                Handle = new WindowInteropHelper(wpfWindow).Handle;
            }

            public IntPtr Handle { get; }
        }
    }
}