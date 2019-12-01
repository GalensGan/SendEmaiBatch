using log4net;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace SendMailBatch
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : INotifyPropertyChanged
    {
        #region 字段
        private readonly string _secretKey = string.Empty;

        private const int SmtpPort = 25;
        private string _fileName;
        private int _totalCount;
        private int _totalNum;
        private int _successNum;
        private int _jumptNum;
        private int _senderIndex = 0;

        private string _mailTitle = "邮件标题";
        private string _mailBody;
        private string _sendInterval = "5";
        private string _selectedSheet;

        private readonly ILog _logger = LogManager.GetLogger(typeof(MainWindow));
        private ObservableCollection<string> _sheetList = new ObservableCollection<string>();
        #endregion

        #region 线程变量
        private string _headerHtml;

        //发件内容
        private readonly ConcurrentQueue<KeyValuePair<string, string>> _queue = new ConcurrentQueue<KeyValuePair<string, string>>();
        //发件人
        private List<EmailAccount> _sender = null;


        private Thread _t;
        private Timer _timer;

        private delegate void AppendTextDelegate(string text);
        #endregion

        #region 属性
        /// <summary>
        /// 邮件标题
        /// </summary>
        public string MailTitle
        {
            get => _mailTitle;
            set
            {
                _mailTitle = value;
                OnPropertyChanged("MailTitle");
            }
        }
        /// <summary>
        /// 邮件正文
        /// </summary>
        public string MailBody
        {
            get => _mailBody;
            set
            {
                _mailBody = value;
                OnPropertyChanged("MailBody");
            }
        }
        /// <summary>
        /// 发送时间间隔
        /// </summary>
        public string SendInterval
        {
            get => _sendInterval;
            set
            {
                _sendInterval = value;
                OnPropertyChanged("SendInterval");
            }
        }

        /// <summary>
        /// 选择的页签
        /// </summary>
        public string SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                _selectedSheet = value;
                OnPropertyChanged("SelectedSheet");
            }
        }
        /// <summary>
        /// Excel页签列表
        /// </summary>
        public ObservableCollection<string> SheetList
        {
            get => _sheetList;
            set
            {
                _sheetList = value;
                OnPropertyChanged("SheetList");
            }
        }

        public Dictionary<string, Employee> EmployeeDic = new Dictionary<string, Employee>();
        #endregion

        #region 构造函数
        public MainWindow(string secretKey)
        {
            InitializeComponent();

            Init();
            DataContext = this;

            if (string.IsNullOrEmpty(secretKey)) _secretKey = "DefaultKey";
            else _secretKey = secretKey;
        }
        #endregion

        #region 初始化
        /// <summary>
        /// 初始化
        /// </summary>
        private void Init()
        {
            //从配置中加载原来的信息            
            MailTitle = EmailConfigManager.EmailTitle;
            MailBody = EmailConfigManager.EmailBody;
            SendInterval = EmailConfigManager.EmailSendIntervalTime;

            XBtnOpen.Click += OnBtnOpen_Click;
            XBtnSend.Click += OnBtnSend_Click;
            XBtnEmployee.Click += OnBtnEmployee_Click;
            XBtnAddSenders.Click += XBtnAddSenders_Click;
        }

        #endregion

        #region 按钮操作
        //添加发件箱
        private void XBtnAddSenders_Click(object sender, RoutedEventArgs e)
        {
            EmailSettingsWpf f = new EmailSettingsWpf(_secretKey);
            f.Owner = this;
            f.ShowDialog();
        }

        /// <summary>
        /// 打开选择文件对话框
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*"
            };
            if (openDlg.ShowDialog().GetValueOrDefault())
            {
                try
                {
                    SelectedSheet = "";
                    SheetList.Clear();
                    using (Stream stream = openDlg.OpenFile())
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook(stream);
                        for (int i = 0; i < workbook.Count; i++)
                        {
                            SheetList.Add(workbook.GetSheetName(i));
                        }
                        if (SheetList.Count > 0)
                        {
                            SelectedSheet = SheetList[0];
                        }
                        stream.Close();
                    }

                    _fileName = openDlg.FileName;
                    XTxtFileName.Text = _fileName;
                }
                catch (IOException ex)
                {
                    _logger.Error(ex.Message, ex);
                    MessageBox.Show(ex.Message);
                }
            }
        }

        /// <summary>
        /// 检查是否满足发送条件
        /// </summary>
        private bool CheckSend()
        {
            if (string.IsNullOrWhiteSpace(_fileName))
            {
                MessageBox.Show("请选择需要导入的Excel文件！");
                return false;
            }
            if (string.IsNullOrWhiteSpace(SelectedSheet))
            {
                MessageBox.Show("请选择需要导入的页签！");
                return false;
            }
            if (string.IsNullOrWhiteSpace(MailTitle))
            {
                MessageBox.Show("请输入邮件标题！");
                return false;
            }
            if (string.IsNullOrWhiteSpace(SendInterval))
            {
                MessageBox.Show("请输入发送时间间隔\n建议大于5秒！");
                return false;
            }

            //读取发件人
            _sender = EmailConfigManager.EmailCounts;
            if (_sender.Count == 0)
            {
                MessageBox.Show("请添加发件箱地址");
                return false;
            }

            return true;
        }

        /// <summary>
        /// 执行发送操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnSend_Click(object sender, RoutedEventArgs e)
        {
            if (CheckSend())
            {
                _totalNum = 0;
                _successNum = 0;
                _headerHtml = "";
                _senderIndex = 0;
                _jumptNum = 0;
                //清空队列
                while (_queue.Count > 0) _queue.TryDequeue(out KeyValuePair<string, string> kv);

                XTxtLog.Text = "";
                XTbkInfo.Text = "";

                // 读取收件人员信息
                ReadEmployees();

                if (EmployeeDic.Count == 0)
                {
                    MessageBox.Show("没有加载到人员邮箱信息，或保存的信息文件不存在！");
                    return;
                }

                if (EmployeeDic.Where(kv => !kv.Value.SendState.Contains("已送达")).Count() == 0)
                {
                    MessageBox.Show("所有人员已经成功发送，如果需要重新发送，请在“添加人员”中清除发送状态");
                    return;
                }

                // 启动线程
                ShowLogInfo("启动线程...", true);
                _t = new Thread(ReadXlsx)
                {
                    IsBackground = true
                };
                _t.Start();

                SendProgress.Visibility = Visibility.Visible;
                XBtnSend.IsEnabled = false;

                //开始发送邮件
                int.TryParse(SendInterval, out int result);
                _timer = new Timer(SendMail, null, 100, result*1000);
                ShowLogInfo("开始发送邮件...");
            }
        }

        /// <summary>
        /// 打开人员信息窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnEmployee_Click(object sender, RoutedEventArgs e)
        {
            EmployeeWindow employeeWin = new EmployeeWindow()
            {
                Owner = this
            };
            employeeWin.ShowDialog();
        }

        #endregion

        #region 读取文件
        /// <summary>
        /// 读取人员信息
        /// </summary>
        private void ReadEmployees()
        {
            EmployeeDic.Clear();

            if (!File.Exists("EmailInfo.xml"))
                return;

            EmailConfigManager.Employees.ForEach(item => EmployeeDic.Add(item.Name, item));
        }

        /// <summary>
        /// 读取文件线程
        /// </summary>
        private void ReadXlsx()
        {
            try
            {
                using (Stream fs = new FileStream(_fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    IWorkbook workbook = WorkbookFactory.Create(fs);
                    ISheet sheet = workbook.GetSheet(SelectedSheet);

                    int firstNum = sheet.FirstRowNum;
                    int lastNum = sheet.LastRowNum;
                    _totalCount = lastNum - firstNum;

                    _headerHtml = "<tr>";
                    IRow headerRow = sheet.GetRow(firstNum);
                    // 设置标题行
                    for (int m = 0; m < headerRow.Cells.Count; m++)
                    {
                        ICell cell = headerRow.Cells[m];
                        string cellVal = ReadCellValue(m, cell);
                        int width = (sheet.GetColumnWidth(m) / 256) * 32;//字符*每个字符的宽度32=16*2                 
                        _headerHtml += string.Format("<th style='border:1px solid black; font-weight:normal;width:{0}px'>{1}</th>", width, cellVal);
                    }
                    _headerHtml += "</tr>";

                    // 设置数据行
                    for (int i = firstNum + 1; i <= lastNum; i++)
                    {
                        string nameIndex = "";
                        string dataHtml = "<tr>";
                        IRow dataRow = sheet.GetRow(i);
                        for (int j = 0; j < dataRow.Cells.Count; j++)
                        {
                            ICell cell = dataRow.Cells[j];
                            string cellVal = ReadCellValue(j, cell);
                            //第2列:人名，为唯一编码
                            if (j == 1) nameIndex = cellVal;
                            dataHtml += string.Format("<td style='border:1px solid black;'>{0}</td>", cellVal);
                        }
                        dataHtml += "</tr>";
                        _queue.Enqueue(new KeyValuePair<string, string>(nameIndex, dataHtml));
                    }

                    fs.Close();
                }
            }
            catch (IOException io)
            {
                _logger.Error(io.Message, io);
                MessageBox.Show("IO异常信息：" + io.Message);
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message, ex);
                MessageBox.Show("读取文件异常信息：" + ex.Message);
            }
        }

        /// <summary>
        /// 读取单元格的值
        /// </summary>
        /// <param name="index">单元格的索引</param>
        /// <param name="cell">指定单元格</param>
        /// <returns>返回值</returns>
        private string ReadCellValue(int index, ICell cell)
        {
            string cellValue;
            switch (cell.CellType)
            {
                case CellType.String:
                    cellValue = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                    cellValue = index > 4
                        ? cell.NumericCellValue.ToString()
                        : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    break;
                case CellType.Boolean:
                    cellValue = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    cellValue = cell.ErrorCellValue.ToString();
                    break;
                default:
                    cellValue = "";
                    break;
            }
            return cellValue.Trim();
        }
        #endregion

        #region 发送邮件
        /// <summary>
        /// 发送邮件线程
        /// </summary>
        private void SendMail(object obj)
        {
            try
            {
                string name = string.Empty;
                string dataHtml = string.Empty;
                if (_queue.Count > 0)
                {
                    if (_queue.TryDequeue(out KeyValuePair<string, string> result))
                    {
                        name = result.Key;
                        dataHtml = result.Value;
                    }
                    else return;
                }
                else if (!_t.IsAlive)//读取线程没有运行了，且队列中没有了数据
                {
                    _timer.Dispose();
                }

                if (string.IsNullOrWhiteSpace(_headerHtml)
                    || string.IsNullOrWhiteSpace(dataHtml))
                    return;

                string addr = string.Empty;
                if (EmployeeDic.TryGetValue(name, out Employee value))
                {
                    if (value.SendState.Contains("已送达"))
                    {
                        _jumptNum++;
                        ShowLogInfo(string.Format("[{0}] 已经发送，本次跳过！", name));
                        if ((_totalNum + _jumptNum) == _totalCount)
                        {
                            //发送完成
                            MarkEnd();
                        }
                        return;
                    }
                    addr = value.Email;
                }

                if (string.IsNullOrWhiteSpace(addr))
                {
                    _totalNum++;
                    ShowLogInfo(string.Format("[{0}] 的邮箱为空，不能发送！！！", name));
                    return;
                }

                string sendMsg =
                    string.Format(
                        @"<p>{0}，您好！</p><p style='margin-left:20px;'>{1}</p>
                        <table style='margin-left:20px; min-width:1500px; border:1px solid black; border-collapse:collapse; 
                               font-size:11pt; text-align:center;'>{2}{3}</table>",
                        name, MailBody, _headerHtml, dataHtml);

                ShowLogInfo(string.Format("开始发送 [{0}]，邮箱：{1}", name, addr));

                if (_senderIndex > _sender.Count - 1) _senderIndex = 0;

                MailAddress fromAddr = new MailAddress(_sender[_senderIndex].AccountName);
                MailAddress toAddr = new MailAddress(addr, addr);
                MailMessage mailMsg = new MailMessage(fromAddr, toAddr)
                {
                    Subject = MailTitle,
                    Body = sendMsg,
                    IsBodyHtml = true,
                    BodyEncoding = Encoding.UTF8
                };

                SmtpClient client = new SmtpClient
                {
                    Host = _sender[_senderIndex].SMTPHostName,
                    Port = SmtpPort,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential(_sender[_senderIndex].AccountName, _sender[_senderIndex].RealPassWord(_secretKey))
                };
                client.SendCompleted += OnSmtpClient_SendCompleted;
                client.SendAsync(mailMsg, name);
                _senderIndex++;
            }
            catch (SmtpException smtp)
            {
                _logger.Error(smtp.Message, smtp);
                MessageBox.Show("SMTP异常信息：" + smtp.Message);
                ShowLogInfo("发送结束。");
                RecoverDefaultState();
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message, ex);
                MessageBox.Show("发送邮件异常信息：" + ex.Message);
                ShowLogInfo("发送结束。");
                RecoverDefaultState();
            }
        }

        private void RecoverDefaultState()
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                XBtnSend.IsEnabled = true;
                SendProgress.Visibility = Visibility.Hidden;
            }));
        }

        /// <summary>
        /// 邮件发送完成事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnSmtpClient_SendCompleted(object sender, AsyncCompletedEventArgs e)
        {
            string state = "";
            if (e.UserState != null)
            {
                state = e.UserState.ToString();
            }

            if (e.Error == null)
            {
                _successNum++;
                ShowLogInfo(string.Format("[{0}] 发送成功！", state));
                EmployeeDic[state].SendState = "已送达";
                EmployeeDic[state].SendDate = DateTime.Now.ToString();
            }
            else
            {
                ShowLogInfo(string.Format("[{0}] 发送失败：{1}", state, e.Error.Message));
                EmployeeDic[state].SendState = "失败";
                EmployeeDic[state].SendDate = DateTime.Now.ToString();
            }

            if ((++_totalNum + _jumptNum) == _totalCount)
            {
                MarkEnd();
                return;
            }
            SetProgress(_totalNum);
        }

        private void MarkEnd()
        {
            SaveConfig();
            ShowLogInfo("发送结束。");
            ShowStatusInfo(string.Format("共发送{0}条数据：成功{1}条，失败{2}条,跳过{3}条。",
                _totalNum, _successNum, _totalNum - _successNum, _jumptNum));
            SetProgress(_totalCount);
        }

        #endregion

        #region 输出日志
        /// <summary>
        /// 输出日志信息
        /// </summary>
        private void ShowLogInfo(string text, bool isFirst = false)
        {
            // 写入日志文件
            _logger.Info(text);

            if (!isFirst)
            {
                text = "\r\n" + text;
            }
            // 输出到界面
            if (Dispatcher.Thread != Thread.CurrentThread)
            {
                Dispatcher.Invoke(new AppendTextDelegate(AppendText), text);
            }
            else
            {
                AppendText(text);
            }
        }
        private void ShowStatusInfo(string text)
        {
            // 写入日志文件
            _logger.Info(text);

            // 输出到界面
            if (Dispatcher.Thread != Thread.CurrentThread)
            {
                Dispatcher.Invoke(new AppendTextDelegate(StatusInfo), text);
            }
            else
            {
                StatusInfo(text);
            }
        }
        private void AppendText(string text)
        {
            XTxtLog.AppendText(text);
            XTxtLog.ScrollToEnd();
        }
        private void StatusInfo(string text)
        {
            XTbkInfo.Text = text;
        }

        /// <summary>
        /// 设置进度条的值
        /// </summary>
        /// <param name="now"></param>
        /// <param name="max"></param>
        private void SetProgress(double now)
        {
            Dispatcher.Invoke(new Action(() =>
            {
               SendProgress.Maximum = _totalCount;
                if (_totalNum+_jumptNum == _totalCount)
                {
                    SendProgress.Visibility = Visibility.Hidden;
                    XBtnSend.IsEnabled = true;
                }
                now = Math.Min(SendProgress.Maximum, now);
            }));           
            Dispatcher.Invoke(new Action<DependencyProperty, object>(SendProgress.SetValue),
                    DispatcherPriority.Background,
                    new object[] { ProgressBar.ValueProperty, now });
        }
        #endregion

        #region 保存配置
        /// <summary>
        /// 保存配置文件
        /// </summary>
        private void SaveConfig()
        {
            //保存邮箱通用设置
            EmailConfigManager.SaveEmailCommonInfo(MailTitle, MailBody, SendInterval);
            //保存发送状态
            EmailConfigManager.SaveEmployees(EmployeeDic.Values.ToList());
        }
        #endregion

        #region 变更通知
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string pPropertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(pPropertyName));
        }
        #endregion
    }
}
