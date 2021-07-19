

#region Using Directives

using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using ExcelExporterImporter.Annotations;

#endregion

namespace ExcelExporterImporter.ViewModels
{
    public class ProgressViewModel : INotifyPropertyChanged
    {
        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        /// <summary>
        ///     Event to notify the owner that the Cancel button was clicked
        /// </summary>
        public event EventHandler<EventArgs> ProcessCanceled;

        #region Constructor/Destructor

        /// <summary>
        ///     Class Constructor
        /// </summary>
        public ProgressViewModel()
        {
            ButtonText = "Cancel";
            Command = new DelegateCommand<object>(OnSubmit, CanSubmit);
        }

        public string ButtonText
        {
            get => buttonText;
            set
            {
                if (value != buttonText) buttonText = value;

                OnPropertyChanged();
            }
        }

        #endregion

        #region Private Fields

        private int _maxValue;
        private string _status;
        private int _value;
        private bool cancelling;
        private string buttonText;

        #endregion

        #region Properties

        /// <summary>
        ///     Used to Bind Buttons elements
        /// </summary>
        public ICommand Command { get; }

        /// <summary>
        ///     Bind to Progress.MaxValue
        /// </summary>
        public int MaxValue
        {
            get => _maxValue;
            set
            {
                if (value == _maxValue) return;
                _maxValue = value;
                OnPropertyChanged();
            }
        }


        /// <summary>
        ///     Bind to Progress.PValue
        /// </summary>
        public int Value
        {
            get => _value;
            set
            {
                if (value == _value) return;
                _value = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        ///     Bind to Status Label
        /// </summary>
        public string Status
        {
            get => _status;
            set
            {
                if (value == _status) return;
                _status = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Methods

        /// <summary>
        ///     ProcessCanceled event invocator
        /// </summary>
        protected virtual void OnProcessCanceled()
        {
            var handler = ProcessCanceled;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public bool ButtonEnabled => !cancelling;

        /// <summary>
        ///     Handle the Click event of the bind buttons
        /// </summary>
        /// <param name="arg"></param>
        private void OnSubmit(object arg)
        {
            if (MessageBox.Show(Resources.AreYouSureYouWantCancelProcess, Resources.Attention, MessageBoxButton.YesNo,
                MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                ButtonText = Resources.Cancelling;
                cancelling = true;
                OnPropertyChanged("ButtonEnabled");
                OnProcessCanceled();
            }
        }

        /// <summary>
        ///     Not used - This is part of ICommand implementation can be used to control
        ///     the buttons click behavior (Can fire the click event or not)
        /// </summary>
        /// <param name="arg"></param>
        /// <returns></returns>
        private bool CanSubmit(object arg)
        {
            return !cancelling;
        }

        /// <summary>
        ///     Start the progress and reset it
        /// </summary>
        /// <param name="maxValue"></param>
        public void Start(int maxValue)
        {
            Value = 0;
            Status = "";
            MaxValue = maxValue;
        }

        /// <summary>
        ///     Increment the progress value
        /// </summary>
        /// <param name="value"></param>
        public void Increment(int value)
        {
            Value += value;
        }

        /// <summary>
        ///     Set the status text
        /// </summary>
        /// <param name="status"></param>
        public void SetStatus(string status)
        {
            Status = status;
        }

        /// <summary>
        ///     End the progress by setting its value to MaxValue
        /// </summary>
        public void End()
        {
            Value = MaxValue;
        }


        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}