using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendMailBatch
{
   public class EmailAccount: INotifyPropertyChanged
    {
        private string _AccountName = string.Empty;
        public string AccountName { get => _AccountName; set { _AccountName = value;OnPropertyChanged("AccountName"); } }

        private string _passWord = string.Empty;
        public string Password { get => _passWord; set { _passWord = value;OnPropertyChanged("PassWord"); } }

        private string _realPassWord = string.Empty;
        public string RealPassWord(string key)
        {
            if (string.IsNullOrEmpty(_realPassWord))
            {
                string md5Key = Encryption.Get32MD5One(key);
                _realPassWord = Encryption.DESDecrypt(Password, md5Key.Substring(0, 8), md5Key.Substring(8, 8));
            }

            return _realPassWord;
        }

        private string _smtpHostName = string.Empty;
        public string SMTPHostName { get => _smtpHostName; set { _smtpHostName = value;OnPropertyChanged("SMTPHostName"); } }

        private bool _isChecked;
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
            AccountName = string.Empty;
            Password = string.Empty;
            SMTPHostName = string.Empty;
        }

        public override bool Equals(object obj)
        {
            if (obj != null && obj is EmailAccount eb && eb.AccountName == this.AccountName) return true;
            else return false;
        }

        public override int GetHashCode()
        {
            return this.AccountName.GetHashCode();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string pPropertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(pPropertyName));
        }
    }
}
