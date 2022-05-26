using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace SubmissionCollector.ViewModel
{
    public abstract class ViewModelBase : INotifyPropertyChanged
    {
        bool _notificationsEnabled = true;
        public event PropertyChangedEventHandler PropertyChanged;

        public void DisableNotifications()
        {
            _notificationsEnabled = false;
        }

        public void EnableNotifications()
        {
            _notificationsEnabled = true;
        }

        public void VerifyPropertyName(string propertyName)
        {
            if (TypeDescriptor.GetProperties(this)[propertyName] == null)
            {
                Debug.Fail("Invalid property name: " + propertyName);
            }
        }

        internal void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            VerifyPropertyName(propertyName);

            if (_notificationsEnabled)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
