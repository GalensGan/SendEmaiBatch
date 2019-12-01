using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml;
using log4net;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SendMailBatch
{
    /// <summary>
    /// EmployeeWindow.xaml 的交互逻辑
    /// </summary>
    public partial class EmailSettingsWpf : Window, INotifyPropertyChanged
    {
        #region 字段
        private readonly string _secretKey = string.Empty;

        private readonly ILog _logger = LogManager.GetLogger(typeof(EmailSettingsWpf));
        private bool _isModify;

        private EmailAccount _addEmailCount = new EmailAccount();
        private ObservableCollection<EmailAccount> _emailCounts = new ObservableCollection<EmailAccount>();
        #endregion

        #region 属性
        
        /// <summary>
        /// 添加账户信息
        /// </summary>
        public EmailAccount AddEmailAccount
        {
            get { return _addEmailCount; }
            set
            {
                _addEmailCount = value;
                OnPropertyChanged("EmailCount");
            }
        }
        /// <summary>
        /// 账户列表
        /// </summary>
        public ObservableCollection<EmailAccount> EmailAccountList
        {
            get { return _emailCounts; }
            set
            {
                _emailCounts = value;
                OnPropertyChanged("EmailCountList");
            }
        }
        #endregion

        #region 构造函数
        public EmailSettingsWpf(string secrectKey)
        {
            InitializeComponent();

            Init();
            DataContext = this;

            _secretKey = secrectKey;
        }
        #endregion

        #region 初始化
        /// <summary>
        /// 初始化
        /// </summary>
        private void Init()
        {
            _isModify = false;
            InitEmailAccount();

            XBtnAdd.Click += OnBtnAdd_Click;
            XBtnDelete.Click += OnBtnDelete_Click;
            XBtnSave.Click += OnBtnSave_Click;
            XChkBoxTitle.Click += OnChkBoxTitle_Click;

            Closing += OnEmployeeWindow_Closing;
        }

        /// <summary>
        /// 初始化账户信息
        /// </summary>
        private void InitEmailAccount()
        {
            EmailAccountList.Clear();
            EmailConfigManager.EmailCounts.ForEach(item => EmailAccountList.Add(item));
        }
        #endregion

        #region 按钮操作

        /// <summary>
        /// 检查是否满足添加条件
        /// </summary>
        private bool CheckAdd()
        {
            string email = AddEmailAccount.AccountName.Trim();
            string stmp = AddEmailAccount.SMTPHostName.Trim();

            if (string.IsNullOrWhiteSpace(email))
            {
                MessageBox.Show("请输入邮箱地址！");
                return false;
            }
            Regex r = new Regex(@"^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$");
            if (!r.IsMatch(email))
            {
                MessageBox.Show("邮箱地址格式不正确！");
                return false;
            }
            if (EmailAccountList.Contains(AddEmailAccount))
            {
                //修改邮箱
                if (MessageBox.Show("邮箱已经存在，是否覆盖？", "温馨提示", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.Cancel) return false;
                EmailAccount account = EmailAccountList[EmailAccountList.IndexOf(AddEmailAccount)];
                string md5Str = Encryption.Get32MD5One(_secretKey);
                account.Password = Encryption.DESEncrypt(EmailPassWordBox.Password, md5Str.Substring(0, 8), md5Str.Substring(8, 8));
                account.SMTPHostName = AddEmailAccount.SMTPHostName;

                AddEmailAccount.Clear();
                EmailPassWordBox.Password = string.Empty;
                _isModify = true;
                return false;
            }
            return true;
        }

        /// <summary>
        /// 添加人员信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAdd())
            {
                EmailAccount newAccount = new EmailAccount
                {
                    AccountName = AddEmailAccount.AccountName.Trim(),
                    SMTPHostName = AddEmailAccount.SMTPHostName,
                };
                string md5Str = Encryption.Get32MD5One(_secretKey);
                newAccount.Password = Encryption.DESEncrypt(EmailPassWordBox.Password, md5Str.Substring(0, 8), md5Str.Substring(8, 8));
                EmailAccountList.Add(newAccount);              

                AddEmailAccount.Clear();
                EmailPassWordBox.Password = string.Empty;

                _isModify = true;
                if (EmailAccountList.Count > 0)
                    XDataGrid.ScrollIntoView(EmailAccountList[EmailAccountList.Count - 1]);
            }
        }        

        /// <summary>
        /// 删除人员信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (EmailAccountList.Count == 0)
                return;

            bool flag = EmailAccountList.Any(emp => emp.IsChecked);
            if (!flag)
            {
                MessageBox.Show("请选择需要删除的人员！");
                return;
            }
            if (MessageBoxResult.OK == MessageBox.Show("您确定要删除选择的人员信息？", "删除人员信息", MessageBoxButton.OKCancel))
            {
                for (int i = 0; i < EmailAccountList.Count; i++)
                {
                    if (EmailAccountList[i].IsChecked)
                    {
                        EmailAccountList.RemoveAt(i--);
                    }
                }
                _isModify = true;
            }
        }

        /// <summary>
        /// 保存人员信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (EmailConfigManager.SaveEmailAccountInfo(EmailAccountList))
            {
                MessageBox.Show("保存成功");
                _isModify = false;
            }
        }

        /// <summary>
        /// 点击全选/全不选
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnChkBoxTitle_Click(object sender, RoutedEventArgs e)
        {
            var chkBox = sender as CheckBox;
            if (chkBox == null || EmailAccountList.Count == 0)
                return;

            foreach (var emp in EmailAccountList)
            {
                emp.IsChecked = chkBox.IsChecked.GetValueOrDefault();
            }
        }
        #endregion

        /// <summary>
        /// 关闭窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnEmployeeWindow_Closing(object sender, CancelEventArgs e)
        {
            if (_isModify && MessageBoxResult.Cancel ==
                MessageBox.Show("系统检测到人员信息已经修改，关闭窗口会放弃修改内容，您确定继续吗？", "", MessageBoxButton.OKCancel))
            {
                e.Cancel = true;
            }
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
