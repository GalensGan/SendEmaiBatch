using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows;
using System.Configuration;
using System.Globalization;
using System.Threading;
using System.Xml;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using log4net;

namespace SendMailDemo
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : INotifyPropertyChanged
    {
        #region 字段
        private const int SmtpPort = 25;
        private string _fileName;
        private int _totalCount;
        private int _totalNum;
        private int _successNum;

        private string _sender;
        private string _mailTitle;
        private string _mailBody;
        private string _smtpHost;
        private string _selectedSheet;

        private readonly ILog _logger = LogManager.GetLogger(typeof(MainWindow));
        private ObservableCollection<string> _sheetList = new ObservableCollection<string>();
        #endregion

        #region 线程变量
        /// <summary>
        /// 同步锁
        /// </summary>
        private static readonly object SyncLock = new Object();

        private string _headerHtml;
        private readonly Queue<string> _queue = new Queue<string>();
        private readonly Queue<string> _nameQueue = new Queue<string>();

        private Thread _t;
        private Timer _timer;

        private delegate void AppendTextDelegate(string text);
        #endregion

        #region 属性
        /// <summary>
        /// 发件人
        /// </summary>
        public string Sender
        {
            get { return _sender; }
            set
            {
                _sender = value;
                OnPropertyChanged("Sender");
            }
        }
        /// <summary>
        /// 邮件标题
        /// </summary>
        public string MailTitle
        {
            get { return _mailTitle; }
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
            get { return _mailBody; }
            set
            {
                _mailBody = value;
                OnPropertyChanged("MailBody");
            }
        }
        /// <summary>
        /// SMTP主机名
        /// </summary>
        public string SmtpHost
        {
            get { return _smtpHost; }
            set
            {
                _smtpHost = value;
                OnPropertyChanged("SmtpHost");
            }
        }

        /// <summary>
        /// 选择的页签
        /// </summary>
        public string SelectedSheet
        {
            get { return _selectedSheet; }
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
            get { return _sheetList; }
            set
            {
                _sheetList = value;
                OnPropertyChanged("SheetList");
            }
        }

        public Dictionary<string, string> EmployeeDic = new Dictionary<string, string>();
        #endregion

        #region 构造函数
        public MainWindow()
        {
            InitializeComponent();

            Init();
            DataContext = this;
        }
        #endregion

        #region 初始化
        /// <summary>
        /// 初始化
        /// </summary>
        private void Init()
        {
            Sender = ConfigurationManager.AppSettings["Sender"];
            XPwdBox.Password = ConfigurationManager.AppSettings["Password"];
            MailTitle = ConfigurationManager.AppSettings["MailTitle"];
            MailBody = ConfigurationManager.AppSettings["MailBody"];
            SmtpHost = ConfigurationManager.AppSettings["SmtpHost"];

            XBtnOpen.Click += OnBtnOpen_Click;
            XBtnSend.Click += OnBtnSend_Click;
            XBtnEmployee.Click += OnBtnEmployee_Click;
        }
        #endregion

        #region 按钮操作
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
            if (string.IsNullOrWhiteSpace(Sender))
            {
                MessageBox.Show("请输入发件人！");
                return false;
            }
            if (string.IsNullOrWhiteSpace(XPwdBox.Password))
            {
                MessageBox.Show("请输入邮箱密码！");
                return false;
            }
            if (string.IsNullOrWhiteSpace(MailTitle))
            {
                MessageBox.Show("请输入邮件标题！");
                return false;
            }
            if (string.IsNullOrWhiteSpace(SmtpHost))
            {
                MessageBox.Show("请输入SMTP主机名！");
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
                lock (SyncLock)
                {
                    _queue.Clear();
                    _nameQueue.Clear();
                }
                XTxtLog.Text = "";
                XTbkInfo.Text = "";

                // 读取人员信息
                ReadEmployees();

                if (EmployeeDic.Count == 0)
                {
                    MessageBox.Show("没有加载到人员邮箱信息，或文件不存在！");
                    return;
                }

                // 启动线程
                ShowLogInfo("启动线程...", true);
                _t = new Thread(ReadXlsx)
                {
                    IsBackground = true
                };
                _t.Start();
                _timer = new Timer(SendMail, null, 100, 500);
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
            try
            {
                if (!File.Exists("EmployeeInfo.xml"))
                    return;

                XmlDocument xml = new XmlDocument();
                xml.Load("EmployeeInfo.xml");

                XmlElement root = xml.DocumentElement;
                if (root == null)
                    return;

                foreach (XmlNode node in root.ChildNodes)
                {
                    var emp = new Employee();
                    foreach (XmlNode node2 in node.ChildNodes)
                    {
                        switch (node2.Name)
                        {
                            case "Name":
                                emp.Name = node2.InnerText;
                                break;
                            case "Email":
                                emp.Email = node2.InnerText;
                                break;
                        }
                    }
                    if (!EmployeeDic.ContainsKey(emp.Name))
                        EmployeeDic.Add(emp.Name, emp.Email);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message, ex);
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// 读取文件线程
        /// </summary>
        private void ReadXlsx()
        {
            try
            {
                using (Stream fs = new FileStream(_fileName, FileMode.Open, FileAccess.Read))
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
                        var cellVal = ReadCellValue(m, cell);
                        _headerHtml += string.Format("<th style='border:1px solid black; font-weight:normal;'>{0}</th>", cellVal);
                    }
                    _headerHtml += "</tr>";

                    // 设置数据行
                    for (int i = firstNum + 1; i <= lastNum; i++)
                    {
                        string name = "";
                        string dataHtml = "<tr>";
                        IRow dataRow = sheet.GetRow(i);
                        for (int j = 0; j < dataRow.Cells.Count; j++)
                        {
                            ICell cell = dataRow.Cells[j];
                            var cellVal = ReadCellValue(j, cell);

                            int width = 80;
                            switch (j)
                            {
                                case 2:  // 姓名列
                                    name = cellVal;
                                    break;
                                case 3:  // 身份证列，较宽
                                    width = 160;
                                    break;
                                default:
                                    width = 80;
                                    break;
                            }
                            dataHtml += string.Format("<td style='border:1px solid black; width:{0}px'>{1}</td>", width, cellVal);
                        }
                        dataHtml += "</tr>";

                        lock (SyncLock)
                        {
                            _queue.Enqueue(dataHtml);
                            _nameQueue.Enqueue(name);
                        }
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
                        ? cell.NumericCellValue.ToString("F2")
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
                string name = "";
                string dataHtml = "";
                if (_queue.Count > 0)
                {
                    lock (SyncLock)
                    {
                        dataHtml = _queue.Dequeue();
                        name = _nameQueue.Dequeue();
                    }
                }
                else
                {
                    if (!_t.IsAlive && _totalNum == _totalCount)
                    {
                        _timer.Dispose();
                    }
                }

                if (string.IsNullOrWhiteSpace(_headerHtml)
                    || string.IsNullOrWhiteSpace(dataHtml))
                    return;

                string addr = "";
                if (EmployeeDic.ContainsKey(name))
                    addr = EmployeeDic[name];
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

                MailAddress fromAddr = new MailAddress(Sender);
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
                    Host = SmtpHost,
                    Port = SmtpPort,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential(Sender, XPwdBox.Password)
                };
                client.SendCompleted += OnSmtpClient_SendCompleted;
                client.SendAsync(mailMsg, name);
            }
            catch (SmtpException smtp)
            {
                _logger.Error(smtp.Message, smtp);
                MessageBox.Show("SMTP异常信息：" + smtp.Message);
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message, ex);
                MessageBox.Show("发送邮件异常信息：" + ex.Message);
            }
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
            }
            else
            {
                ShowLogInfo(string.Format("[{0}] 发送失败：{1}", state, e.Error.Message));
            }

            if (++_totalNum == _totalCount)
            {
                SaveConfig();
                ShowLogInfo("发送结束。");
                ShowStatusInfo(string.Format("共发送{0}条数据：成功{1}条，失败{2}条。",
                    _totalNum, _successNum, _totalNum - _successNum));
            }
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
        #endregion

        #region 保存配置
        /// <summary>
        /// 保存配置文件
        /// </summary>
        private void SaveConfig()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["Sender"].Value = Sender;
            config.AppSettings.Settings["Password"].Value = XPwdBox.Password;
            config.AppSettings.Settings["MailTitle"].Value = MailTitle;
            config.AppSettings.Settings["MailBody"].Value = MailBody;
            config.AppSettings.Settings["SmtpHost"].Value = SmtpHost;
            config.Save();
        }
        #endregion

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
