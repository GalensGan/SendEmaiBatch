using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendMailDemo
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
            if (obj == null)
                return false;

            if (obj.GetType() != GetType())
                return false;

            return Equals(obj as Employee);
        }
        private bool Equals(Employee emp)
        {
            return (Name == emp.Name || Email == emp.Email);
        }

        public override int GetHashCode()
        {
            return (Name + Email).GetHashCode();
        }

        #region 变更通知
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string pPropertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(pPropertyName));
            }
        }
        #endregion
    }
}
