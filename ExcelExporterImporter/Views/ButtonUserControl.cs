using System.Windows.Forms;
using System.Windows.Input;

namespace ExcelExporterImporter.Views
{
    public partial class ButtonUserControl : UserControl
    {
        public ButtonUserControl()
        {
            InitializeComponent();
            button.Click += (sender, args) => OnButtonClick();
        }

        public string ButtonText
        {
            get => button.Text;
            set => button.Text = value;
        }

        public ICommand Command { get; set; }

        private void OnButtonClick()
        {
            if (Command.CanExecute(null)) Command.Execute(null);
        }
    }
}