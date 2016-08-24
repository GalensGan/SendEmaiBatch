#SendMailForWife
该项目实现了公司发放工资条功能，可以从Excel导入人员及工资信息。
使用的技术手段包括：
1、log4net记录日志
2、NPOI读取Excel信息
3、XmlDocument读取XML文件
4、SmtpClient发送邮件
5、多线程：使用单独的线程读取人员工资信息，将读取的信息进行缓存；利用Timer计时器定时从缓存中获取发送数据，发送完所有文件数据后销毁。
