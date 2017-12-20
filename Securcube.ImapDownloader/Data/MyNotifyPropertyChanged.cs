using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.Serialization;


namespace SecurCube.ImapDownloader.Data
{

    [Serializable]
    abstract class MyNotifyPropertyChanged : INotifyPropertyChanged
    {

        #region INotifyPropertyChanged
        [field: NonSerializedAttribute()]
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetField<T>(ref T field, T value, string propertyName)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
        #endregion

    }
}
