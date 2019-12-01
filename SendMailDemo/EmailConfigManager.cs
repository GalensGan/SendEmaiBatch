using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Xml;

namespace SendMailBatch
{
    internal class EmailConfigManager
    {
        private static readonly string _emailInfoFileName = "EmailInfo.xml";
        /// <summary>
        /// 返回员工信息
        /// </summary>
        public static List<Employee> Employees
        {
            get
            {
                if (!File.Exists(_emailInfoFileName))
                {
                    return new List<Employee>();
                }
                else
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(_emailInfoFileName);
                    XmlNode xmlNode = doc.SelectSingleNode("EmailInfo/Employees");
                    List<Employee> resultList = new List<Employee>();
                    if (xmlNode != null)
                    {
                        foreach (XmlNode node in xmlNode.ChildNodes)
                        {
                            Employee emp = new Employee();
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
                                    case "SendState":
                                        emp.SendState = node2.InnerText;
                                        break;
                                    case "SendDate":
                                        emp.SendDate = node2.InnerText;
                                        break;
                                    default:
                                        break;
                                }
                            }
                            resultList.Add(emp);
                        }
                    }
                    return resultList;
                }
            }
        }

        /// <summary>
        /// 保存收件人信息
        /// </summary>
        /// <param name="employees"></param>
        /// <returns></returns>
        public static bool SaveEmployees(IList<Employee> employees)
        {
            XmlDocument doc = new XmlDocument();
            if (!File.Exists(_emailInfoFileName))
            {
                //声明
                XmlNode node = doc.CreateXmlDeclaration("1.0", "utf-8", "");
                doc.AppendChild(node);
                XmlElement xmlElement = doc.CreateElement("EmailInfo");
                doc.AppendChild(xmlElement);
            }
            else
            {
                doc.Load(_emailInfoFileName);

            }
            XmlNode root = doc.SelectSingleNode("EmailInfo/Employees");
            if (root != null)
            {
                //清空root
                List<XmlNode> childNodes = new List<XmlNode>();
                foreach (XmlNode sub in root.ChildNodes) childNodes.Add(sub);
                childNodes.ForEach(item => root.RemoveChild(item));
            }
            else
            {
                root = doc.CreateElement("Employees");
                doc.DocumentElement.AppendChild(root);
            }
            foreach (Employee emp in employees)
            {
                XmlElement ele = doc.CreateElement("Employee");
                XmlElement ele2 = doc.CreateElement("Name");
                ele2.InnerText = emp.Name;
                ele.AppendChild(ele2);

                ele2 = doc.CreateElement("Email");
                ele2.InnerText = emp.Email;
                ele.AppendChild(ele2);

                ele2 = doc.CreateElement("SendState");
                ele2.InnerText = emp.SendState;
                ele.AppendChild(ele2);

                ele2 = doc.CreateElement("SendDate");
                ele2.InnerText = emp.SendDate;
                ele.AppendChild(ele2);

                if (root != null) root.AppendChild(ele);
            }
            doc.Save(_emailInfoFileName);
            return true;
        }

        /// <summary>
        /// 获取发件箱列表
        /// </summary>
        public static List<EmailAccount> EmailCounts
        {
            get
            {
                if (!File.Exists(_emailInfoFileName))
                {
                    return new List<EmailAccount>();
                }
                else
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(_emailInfoFileName);
                    XmlNode xmlNode = doc.SelectSingleNode("EmailInfo/EmailAccounts");
                    List<EmailAccount> resultList = new List<EmailAccount>();
                    if (xmlNode != null)
                    {
                        foreach (XmlNode node in xmlNode.ChildNodes)
                        {
                            EmailAccount emc = new EmailAccount();
                            foreach (XmlNode node2 in node.ChildNodes)
                            {
                                switch (node2.Name)
                                {
                                    case "AccountName":
                                        emc.AccountName = node2.InnerText;
                                        break;
                                    case "PassWord":
                                        emc.Password = node2.InnerText;
                                        break;
                                    case "SMTPHostName":
                                        emc.SMTPHostName = node2.InnerText;
                                        break;
                                    default:
                                        break;
                                }
                            }
                            resultList.Add(emc);
                        }
                    }
                    return resultList;
                }
            }
        }


        /// <summary>
        /// 保存发件箱
        /// </summary>
        /// <param name="emailCounts"></param>
        /// <returns></returns>
        public static bool SaveEmailAccountInfo(IList<EmailAccount> emailCounts)
        {
            XmlDocument doc = new XmlDocument();
            if (!File.Exists(_emailInfoFileName))
            {
                //声明
                XmlNode node = doc.CreateXmlDeclaration("1.0", "utf-8", "");
                doc.AppendChild(node);
                XmlElement xmlElement = doc.CreateElement("EmailInfo");
                doc.AppendChild(xmlElement);
            }
            else
            {
                doc.Load(_emailInfoFileName);

            }
            XmlNode root = doc.SelectSingleNode("EmailInfo/EmailAccounts");
            if (root != null)
            {
                //清空root
                List<XmlNode> childNodes = new List<XmlNode>();
                foreach (XmlNode sub in root.ChildNodes) childNodes.Add(sub);
                childNodes.ForEach(item => root.RemoveChild(item));
            }
            else
            {
                root = doc.CreateElement("EmailAccounts");
                doc.DocumentElement.AppendChild(root);
            }
            foreach (EmailAccount emc in emailCounts)
            {
                XmlElement ele = doc.CreateElement("EmailAccount");
                XmlElement ele2 = doc.CreateElement("AccountName");
                ele2.InnerText = emc.AccountName;
                ele.AppendChild(ele2);

                ele2 = doc.CreateElement("PassWord");
                ele2.InnerText = emc.Password;
                ele.AppendChild(ele2);

                ele2 = doc.CreateElement("SMTPHostName");
                ele2.InnerText = emc.SMTPHostName;
                ele.AppendChild(ele2);

                if (root != null) root.AppendChild(ele);
            }
            doc.Save(_emailInfoFileName);
            return true;
        }

        public static string EmailTitle
        {
            get
            {
                if (!File.Exists(_emailInfoFileName))
                {
                    return string.Empty;
                }
                else
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(_emailInfoFileName);
                    XmlNode xmlNode = doc.SelectSingleNode("EmailInfo/EmailCommon/EmailTitle");
                    if (xmlNode != null)
                    {
                        return xmlNode.InnerText;
                    }
                    else return string.Empty;
                }
            }
        }

        public static string EmailBody
        {
            get
            {
                if (!File.Exists(_emailInfoFileName))
                {
                    return string.Empty;
                }
                else
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(_emailInfoFileName);
                    XmlNode xmlNode = doc.SelectSingleNode("EmailInfo/EmailCommon/EmailBody");
                    if (xmlNode != null)
                    {
                        return xmlNode.InnerText;
                    }
                    else return string.Empty;
                }
            }
        }

        public static string EmailSendIntervalTime
        {
            get
            {
                if (!File.Exists(_emailInfoFileName))
                {
                    return string.Empty;
                }
                else
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(_emailInfoFileName);
                    XmlNode xmlNode = doc.SelectSingleNode("EmailInfo/EmailCommon/SendInterval");
                    if (xmlNode != null)
                    {
                        return xmlNode.InnerText;
                    }
                    else return string.Empty;
                }
            }
        }

        public static bool SaveEmailCommonInfo(string title, string body, string sendInterval)
        {
            XmlDocument doc = new XmlDocument();
            if (!File.Exists(_emailInfoFileName))
            {
                //声明
                XmlNode node = doc.CreateXmlDeclaration("1.0", "utf-8", "");
                doc.AppendChild(node);
                XmlElement xmlElement = doc.CreateElement("EmailInfo");
                doc.AppendChild(xmlElement);
            }
            else
            {
                doc.Load(_emailInfoFileName);
            }
            XmlNode root = doc.SelectSingleNode("EmailInfo/EmailCommon");
            if (root != null) root.ParentNode.RemoveChild(root);
            root = doc.DocumentElement;

            XmlElement ele = doc.CreateElement("EmailCommon");
            XmlElement ele2 = doc.CreateElement("EmailTitle");
            ele2.InnerText = title;
            ele.AppendChild(ele2);

            ele2 = doc.CreateElement("EmailBody");
            ele2.InnerText = body;
            ele.AppendChild(ele2);

            ele2 = doc.CreateElement("SendInterval");
            ele2.InnerText = sendInterval;
            ele.AppendChild(ele2);

            if (root != null) root.AppendChild(ele);
            doc.Save(_emailInfoFileName);
            return true;
        }
    }
}
