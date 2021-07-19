using System.ComponentModel;

namespace ExcelExporterImporter.Common
{
    public class CheckedListItem<T> : INotifyPropertyChanged
    {
        private string _id;
        private bool _isChecked;
        private T _item;
        private object _object;

        /// <summary>
        ///     Check all the fields of a list
        /// </summary>
        /// <param name="id"></param>
        /// <param name="item"></param>
        /// <param name="obj"></param>
        /// <param name="isChecked"></param>
        public CheckedListItem(string id, T item, object obj = null, bool isChecked = false)
        {
            _item = item;
            _isChecked = isChecked;
            _id = id;
            _object = obj;
        }

        public T Item
        {
            get => _item;
            set
            {
                _item = value;
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("Item"));
            }
        }

        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                _isChecked = value;
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("IsChecked"));
            }
        }

        public string Id
        {
            get => _id;
            set
            {
                _id = value;
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("Id"));
            }
        }

        public object Object
        {
            get => _object;
            set
            {
                _object = value;
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("Object"));
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;
    }
}