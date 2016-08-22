using System;
using System.ComponentModel;

using System.Runtime.CompilerServices;
using System.Linq.Expressions;
using System.Windows;


namespace LumisCalendarSync.ViewModels
{
    public class BindableBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual bool Set<T>(ref T storage, T value, string propertyName)
        {
            if (Equals(storage, value)) return false;

            storage = value;
            RaisePropertyChanged(propertyName);

            return true;
        }

        protected void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private static DependencyObject myDesignDependencyObject = new DependencyObject();
        public static bool InDesignMode()
        {
            return DesignerProperties.GetIsInDesignMode(myDesignDependencyObject);
        }

    }
}
