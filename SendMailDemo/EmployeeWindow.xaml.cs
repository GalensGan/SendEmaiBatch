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

namespace SendMailDemo
{
    /// <summary>
    /// EmployeeWindow.xaml 的交互逻辑
    /// </summary>
    public partial class EmployeeWindow : Window, INotifyPropertyChanged
    {
        #region 字段
        private readonly ILog _logger = LogManager.GetLogger(typeof(EmployeeWindow));

        private string _fileName;
        private string _selectedSheet;
        private bool _isModify;

        private Employee _addEmp = new Employee();
        private ObservableCollection<string> _sheetList = new ObservableCollection<string>();
        private ObservableCollection<Employee> _employeeList = new ObservableCollection<Employee>();
        #endregion

        #region 属性
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
        /// <summary>
        /// 添加人员对象
        /// </summary>
        public Employee AddEmp
        {
            get { return _addEmp; }
            set
            {
                _addEmp = value;
                OnPropertyChanged("AddEmp");
            }
        }
        /// <summary>
        /// 人员列表
        /// </summary>
        public ObservableCollection<Employee> EmployeeList
        {
            get { return _employeeList; }
            set
            {
                _employeeList = value;
                OnPropertyChanged("EmployeeList");
            }
        }
        #endregion

        #region 构造函数
        public EmployeeWindow()
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
            _isModify = false;
            InitEmployees();

            XBtnAdd.Click += OnBtnAdd_Click;
            XBtnOpen.Click += OnBtnOpen_Click;
            XBtnImport.Click += OnBtnImport_Click;
            XBtnDelete.Click += OnBtnDelete_Click;
            XBtnSave.Click += OnBtnSave_Click;
            XChkBoxTitle.Click += OnChkBoxTitle_Click;

            Closing += OnEmployeeWindow_Closing;
        }

        /// <summary>
        /// 初始化人员信息
        /// </summary>
        private void InitEmployees()
        {
            EmployeeList.Clear();
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
                    EmployeeList.Add(emp);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message, ex);
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 按钮操作

        /// <summary>
        /// 检查是否满足添加条件
        /// </summary>
        private bool CheckAdd()
        {
            string name = AddEmp.Name.Trim();
            string email = AddEmp.Email.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("请输入人员姓名！");
                return false;
            }
            if (string.IsNullOrWhiteSpace(email))
            {
                MessageBox.Show("请输入邮箱地址！");
                return false;
            }
            Regex r = new Regex(@"^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$");
            if (!r.IsMatch(AddEmp.Email))
            {
                MessageBox.Show("邮箱地址格式不正确！");
                return false;
            }
            if (EmployeeList.Contains(AddEmp))
            {
                MessageBox.Show("姓名或邮箱已经存在，不能再次添加！");
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
                EmployeeList.Add(new Employee
                {
                    Name = AddEmp.Name.Trim(),
                    Email = AddEmp.Email.Trim()
                });
                AddEmp.Clear();

                _isModify = true;
                if (EmployeeList.Count > 0)
                    XDataGrid.ScrollIntoView(EmployeeList[EmployeeList.Count - 1]);
            }
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
        /// 从文件导入人员信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnImport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_fileName))
            {
                MessageBox.Show("请选择需要导入的Excel文件！");
                return;
            }
            if (string.IsNullOrWhiteSpace(SelectedSheet))
            {
                MessageBox.Show("请选择需要导入的页签！");
                return;
            }

            try
            {
                using (Stream fs = new FileStream(_fileName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = WorkbookFactory.Create(fs);
                    ISheet sheet = workbook.GetSheet(SelectedSheet);

                    int firstNum = sheet.FirstRowNum;
                    int lastNum = sheet.LastRowNum;
                    int totalCount = lastNum - firstNum;
                    int successNum = 0;

                    // 设置数据行
                    _logger.Info("开始导入人员信息...");
                    Regex r = new Regex(@"^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$");
                    for (int i = firstNum + 1; i <= lastNum; i++)
                    {
                        Employee emp = new Employee();
                        IRow dataRow = sheet.GetRow(i);
                        for (int j = 0; j < 2; j++)
                        {
                            ICell cell = dataRow.Cells[j];
                            var cellVal = ReadCellValue(j, cell);
                            switch (j)
                            {
                                case 0:  // 姓名
                                    emp.Name = cellVal;
                                    break;
                                case 1:  // 邮箱
                                    emp.Email = cellVal;
                                    break;
                            }
                        }

                        if (string.IsNullOrWhiteSpace(emp.Name) || string.IsNullOrWhiteSpace(emp.Email))
                        {
                            _logger.Warn(string.Format("第[{0}]条记录的姓名或邮箱为空，姓名：{1}，邮箱：{2}", i, emp.Name, emp.Email));
                            continue;
                        }
                        if (!r.IsMatch(emp.Email))
                        {
                            _logger.Warn(string.Format("第[{0}]条记录的邮箱格式不正确，姓名：{1}，邮箱：{2}", i, emp.Name, emp.Email));
                            continue;
                        }
                        if (EmployeeList.Contains(emp))
                        {
                            _logger.Warn(string.Format("原集合中已经包含第[{0}]条记录的姓名或邮箱，姓名：{1}，邮箱：{2}", i, emp.Name, emp.Email));
                            continue;
                        }

                        EmployeeList.Add(emp);
                        successNum++;
                        _logger.Info(string.Format("[{0}] 导入成功！", emp.Name));
                    }

                    _isModify = true;
                    if (EmployeeList.Count > 0)
                        XDataGrid.ScrollIntoView(EmployeeList[EmployeeList.Count - 1]);

                    _logger.Info("导入完成。");
                    string info = string.Format("共导入{0}条数据：成功{1}条，失败{2}条。",
                        totalCount, successNum, totalCount - successNum);
                    _logger.Info(info);
                    MessageBox.Show(info);

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

        /// <summary>
        /// 删除人员信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (EmployeeList.Count == 0)
                return;

            bool flag = EmployeeList.Any(emp => emp.IsChecked);
            if (!flag)
            {
                MessageBox.Show("请选择需要删除的人员！");
                return;
            }
            if (MessageBoxResult.OK == MessageBox.Show("您确定要删除选择的人员信息？", "删除人员信息", MessageBoxButton.OKCancel))
            {
                for (int i = 0; i < EmployeeList.Count; i++)
                {
                    if (EmployeeList[i].IsChecked)
                    {
                        EmployeeList.RemoveAt(i--);
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
            try
            {
                XmlDocument xml = new XmlDocument();
                xml.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?> <Employees></Employees>");

                XmlElement root = xml.DocumentElement;
                foreach (Employee emp in EmployeeList)
                {
                    XmlElement ele = xml.CreateElement("Employee");
                    XmlElement ele2 = xml.CreateElement("Name");
                    ele2.InnerText = emp.Name;
                    ele.AppendChild(ele2);

                    ele2 = xml.CreateElement("Email");
                    ele2.InnerText = emp.Email;
                    ele.AppendChild(ele2);

                    if (root != null) root.AppendChild(ele);
                }
                xml.Save("EmployeeInfo.xml");
                MessageBox.Show("保存成功！");

                _isModify = false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message, ex);
                MessageBox.Show(ex.Message);
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
            if (chkBox == null || EmployeeList.Count == 0)
                return;

            foreach (var emp in EmployeeList)
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
