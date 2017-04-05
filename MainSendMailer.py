# -*- coding: utf-8 -*-
from PyQt5 import QtGui, QtCore,QtWidgets
from MailNeW.sendmail import Ui_SendMailer
import os
# import icoqrc


class SendingMail(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(SendingMail, self).__init__(parent)
        self.Ui= Ui_SendMailer()
        self.Ui.setupUi(self)
        self.setWindowTitle(u'Pyqt 邮件发送')
        # self.setWindowIcon(QtGui.QIcon('myfavicon.ico'))
        # reload(sys)
        # sys.setdefaultencoding("utf-8")
        #初始化UI项目
        self.proparam()
        # 获取配置项
        self.GetConfig()
        self.Ui.comboBoxServer.activated.connect(self.ServerSmtp)   # stmp服务器
        self.Ui.attachBtn.clicked.connect(self.SelAttach)    # 选择附件
        # self.connect(self.Ui.sendingBtn,QtCore.SIGNAL("clicked()"), self.sending)
        self.Ui.sendingBtn.clicked.connect(self.CoreAction)

        # 发送进度
        self.timer = QtCore.QBasicTimer()

    # def初始化UI项目
    def proparam(self):
        self.Ui.comboBoxServer.addItem(u'选择服务地址',QtCore.QVariant(''))
        self.Ui.comboBoxServer.addItem(u'网易163',QtCore.QVariant('smtp.163.com'))
        self.Ui.comboBoxServer.addItem(u'腾讯邮箱',QtCore.QVariant('smtp.qq.com'))
        self.Ui.comboBoxServer.addItem(u'新浪邮箱',QtCore.QVariant('smtp.sina.com.cn'))
        self.Ui.comboBoxServer.addItem(u'搜狐邮箱',QtCore.QVariant('smtp.sohu.com'))
        self.Ui.comboBoxServer.addItem(u'Google邮箱',QtCore.QVariant('smtp.gmail.com'))
        self.Ui.comboBoxServer.addItem(u'自定义',QtCore.QVariant(''))
        # 隐藏发送进度和发送提示
        self.Ui.progressBarSend.hide()
        self.Ui.labelSendMsg.hide()
        # 记住密码配置密钥
        self.keyt=2711790504
        self.file_version =100

    # stmp 服务器
    def ServerSmtp(self):
        comboxserver= self.Ui.comboBoxServer.currentIndex()
        comboxItemData= self.Ui.comboBoxServer.itemData(comboxserver)
        self.Ui.lineEditSMTP.setText(comboxItemData)

    # 选择附件  后期判断文件大小，显示格式等
    def SelAttach(self):
        files = QtWidgets.QFileDialog.getOpenFileName(self, u'选择附件')
        # print(files)
        self.Ui.lineEditAttach.setText(files[0])

    # 配置项
    def GetConfig(self, sendingSucc=''):
        isConfig = os.path.exists('SendMailerConfig.ini')
        # print('isConfig = %s' % isConfig)
        if not isConfig:
            os.makedirs('SendMailerConfig.ini')
        else:
            if sendingSucc == '':  # 存在记住用户名和密码
                self.Ui.checkBoxRememberUP.setChecked(True)
                file = QtCore.QFile("SendMailerConfig.ini")
                # print('file = %s' % file)
                In = QtCore.QDataStream(file)
                In.setVersion(QtCore.QDataStream.Qt_5_8)
                # In.writeInt32(self.keyt)
                # file.open(QtCore.QIODevice.ReadOnly)
                # print('#######')
                if not file.open(QtCore.QIODevice.ReadOnly):
                    raise IOError(str(file.errorString()))
                magic = In.readUInt32()
                # print('magic = %s' % magic)
                if magic != self.keyt:
                    QtWidgets.QMessageBox.information(self,"exception",u"记住密码数据格式错误！")
                    return
                comboBoxServer = In.readUInt32()
                LoginUser = In.readQString()
                LoginPasswd = In.readQString()
                self.Ui.comboBoxServer.setCurrentIndex(comboBoxServer)
                comboxItemData = self.Ui.comboBoxServer.itemData(comboBoxServer)
                # print('comboxItemData =',comboxItemData)
                self.Ui.lineEditSMTP.setText(comboxItemData)
                self.Ui.lineEditUser.setText(LoginUser)
                self.Ui.lineEditPasswd.setText(LoginPasswd)
            # 发送成功判断记住密码
            if sendingSucc:
                isChecked = self.Ui.checkBoxRememberUP.isChecked()
                if isChecked:
                    file = QtCore.QFile("SendMailerConfig.ini")
                    self.time = QtCore.QDateTime()
                    file.open(QtCore.QIODevice.WriteOnly)
                    out = QtCore.QDataStream(file)
                    out.setVersion(QtCore.QDataStream.Qt_5_8)
                    out.writeUInt32(int(self.keyt))  # writeUInt32  参数为int类型
                    out.writeUInt32(self.Ui.comboBoxServer.currentIndex())
                    out.writeQString(self.Ui.lineEditUser.text())  # writeString 参数为string类型
                    out.writeQString(self.Ui.lineEditPasswd.text())
                    out << self.time.currentDateTime()

    # # 发送邮件
    # def sending(self):
    #     exist = self.basicExist()
    #     if exist:
    #         success = self.CoreAction()
    #         if success:
    #             #判断是否记住密码
    #             self.GetConfig('sendSuccess')
    #             self.timer.start(10, self)

    # 引入smtplib 发送邮件
    def CoreAction(self):
        validate = self.basicExist()
        if validate:
            # 发送进度显示
            self.Ui.labelSendMsg.show()
            self.Ui.labelSendMsg.setText(u'正在发送……')
            self.Ui.progressBarSend.show()
            self.Ui.progressBarSend.reset()  # 重置进度条
            self.step = 0
            self.timer.start(1000, self)
            DictData = {}
            smtpConfig = {}  # 配置
            smtpConfig['Connect'] = str(self.Ui.lineEditSMTP.text())
            smtpConfig['LoginUser'] = str(self.Ui.lineEditUser.text())
            smtpConfig['LoginPasswd'] = str(self.Ui.lineEditPasswd.text())

            # 判断Longinuser
            isExist=smtpConfig['LoginUser']
            if '@' not in isExist:
                smtpConfig['LoginUser'] = smtpConfig['LoginUser'] + smtpConfig['Connect'].replace('smtp.', '@', 1)

            DictData['smtpConfig'] = smtpConfig  # 配置
            DictData['Subject'] = str(self.Ui.lineEditCheme.text())  # 主题
            DictData['content'] = str(self.Ui.textEditBody.toPlainText())  # 内容
            DictData['attach'] = str(self.Ui.lineEditAttach.text())  # 附件
            Receiver = str(self.Ui.lineEditReceiver.text())
            DictData['ListReceiver'] = Receiver.split(',')
            self.Theading = Theading(DictData)
            # self.connect(self.Theading, QtCore.SIGNAL("updateresult"), self.updateResult)  # 创建一个信号，在线程状态结果时发射触发
            self.Theading.sinOut.connect(self.updateResult)
            self.Theading.start()  # 线程开始
            # self.disconnect(self.Ui.sendingBtn, QtCore.SIGNAL("clicked()"), self.CoreAction)  # 取消connect事件
            self.Ui.sendingBtn.clicked.disconnect(self.CoreAction)
            # self.Theading.sinOut.disconnect(self.CoreAction)

            self.Ui.sendingBtn.setEnabled(False)

            # 发送邮件后，进程触发事件

    # 发送邮件后，进程触发事件
    def updateResult(self, status):
        if status['status'] == 1:  # 发送成功
            # 判断是否记住密码
            self.GetConfig('sendSuccess')
            self.timer.start(10, self)

        else:  # 发送失败
            self.Ui.labelSendMsg.hide()
            self.Ui.progressBarSend.hide()
            QtWidgets.QMessageBox.warning(self, u'错误提示', status['msg'], QtWidgets.QMessageBox.Yes)
        self.Ui.sendingBtn.clicked.connect(self.CoreAction)
        # self.connect(self.Ui.sendingBtn, QtCore.SIGNAL("clicked()"), self.CoreAction)  # connect事件
        self.Ui.sendingBtn.setEnabled(True)

    '''
     获取配置项目
    '''
    def basicExist(self):
        if self.Ui.lineEditSMTP.text() == '':
            QtWidgets.QMessageBox.warning(self, (u'提示'),(u'请填写SMTP服务地址'), QtWidgets.QMessageBox.Yes)
            return False
        if self.Ui.lineEditUser.text() == '':
            QtWidgets.QMessageBox.warning(self, (u'提示'),(u'请填写发送者用户名'), QtWidgets.QMessageBox.Yes)
            return False
        if self.Ui.lineEditPasswd.text() == '':
            QtWidgets.QMessageBox.warning(self, (u'提示'),(u'请填写发送者密码'), QtWidgets.QMessageBox.Yes)
            return False
        if self.Ui.lineEditReceiver.text() == '':
            QtWidgets.QMessageBox.warning(self, (u'提示'),(u'请填写接收者邮箱'), QtWidgets.QMessageBox.Yes)
            return False
        if self.Ui.lineEditCheme.text() == '':
            QtWidgets.QMessageBox.warning(self, (u'提示'),(u'请填写邮件标题'), QtWidgets.QMessageBox.Yes)
            return False
        return True

    def keyPressEvent(self, event):
        if event.key() ==QtCore.Qt.Key_Escape:
            Ok = QtWidgets.QMessageBox.question(self,u'提示', u'确定要退出吗？',QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.No)
            if Ok == QtWidgets.QMessageBox.Yes:
                self.close()
    # 重写timer
    def timerEvent(self,event):
        if self.step >= 100:
            self.timer.stop()
            self.Ui.labelSendMsg.show()
            self.Ui.labelSendMsg.setText(u"发送完成！")
            self.Ui.sendingBtn.setEnabled(True)
            return
        self.step = self.step+1
        self.Ui.progressBarSend.setValue(self.step)

class Theading(QtCore.QThread):
    sinOut = QtCore.pyqtSignal(dict)
    def __init__(self, dicts, parent=None):
        super(Theading, self).__init__(parent)
        self.dict = dicts
    def run(self):
        # 引入smtplib 发送邮件
        import smtplib
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        result = {}
        result['status'] = 0
        result['msg'] = ''
        try:
            smtp = smtplib.SMTP_SSL()
            smtp.connect(self.dict['smtpConfig']['Connect'])
            smtp.login(self.dict['smtpConfig']['LoginUser'], self.dict['smtpConfig']['LoginPasswd'])
            msg = MIMEMultipart()
            msg['Subject'] = self.dict['Subject']
            msg['From'] = self.dict['smtpConfig']['LoginUser']
            from email.header import Header
            msg['To'] = Header('TA', 'utf-8')
            # 内容
            content = MIMEText(self.dict['content'], 'html', 'utf-8')   # 获取textedit数据
            # 附件
            isExistAttach = self.dict['attach']
            if isExistAttach:  # 如果存在附件
                filename = os.path.basename(isExistAttach)
                attach = MIMEText(open(isExistAttach, 'rb').read(), 'base64', 'gb2312')
                attach["Content-Type"] = 'application/octet-stream'
                attach["Content-Disposition"] = 'attachment; filename="'+filename+'"'
                msg.attach(attach)  # 可能有附件没有内容

            msg.attach(content)
            # Receiver=str(self.Ui.lineEditReceiver.text())
            # print('Receiver = ',Receiver)
            # ListReceiver= Receiver.split(',')
            smtp.sendmail(self.dict['smtpConfig']['LoginUser'], self.dict['ListReceiver'], msg.as_string())  # 发送者 和 接收人
            smtp.quit()
            result['status'] = 1

        except Exception as e:
            result['msg'] = str(e)
            # print(result)
        self.sinOut.emit(result)


if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    appSendingMail = SendingMail()
    appSendingMail.show()
    sys.exit(app.exec_())