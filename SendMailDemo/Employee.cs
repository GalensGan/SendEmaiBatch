using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendMailBatch
{
    public class Employee : INotifyPropertyChanged
    {
        private string _name = "";
        private string _email = "";
        private bool _isChecked;

        /// <summary>
        /// 名称
        /// </summary>
        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                OnPropertyChanged("Name");
            }
        }
        /// <summary>
        /// 邮件地址
        /// </summary>
        public string Email
        {
            get { return _email; }
            set
            {
                _email = value;
                OnPropertyChanged("Email");
            }
        }

        private string _sendState = string.Empty;
        public string SendState
        {
            get => _sendState;
            set
            {
                _sendState = value;
                OnPropertyChanged("SendState");
            }
        }

        private string _sendDate = string.Empty;
        public string SendDate
        {
            get => _sendDate;
            set
            {
                _sendDate = value;
                OnPropertyChanged("SendDate");
            }
        }

        /// <summary>
        /// 是否选中
        /// </summary>
        public bool IsChecked
        {
            get { return _isChecked; }
            set
            {
                _isChecked = value;
                OnPropertyChanged("IsChecked");
            }
        }

        public void Clear()
        {
            Name = string.Empty;
            Email = string.Empty;
            IsChecked = false;
        }

        public override bool Equals(object obj)
        {
            if (obj != null && obj is Employee e && e.Name == this.Name && e.Email == this.Email) return true;
            else return false;
        }

        public override int GetHashCode()
        {
            return (Name + Email).GetHashCode();
        }

        #region 变更通知
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string pPropertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(pPropertyName));
        }
        #endregion
    }
}
