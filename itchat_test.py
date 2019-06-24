import itchat
from excel_CBT import excel_read

def download_files(msg):
    msg.download("./" + msg['FileName'])



# 自动回复
# 封装好的装饰器，当接收到的消息是Text，即文字消息
@itchat.msg_register(['Attachment'])
def text_reply(msg):
    # 当消息不是由自己发出的时候
    if not msg['FromUserName'] == myUserName:
        # 发送一条提示给文件助手
        itchat.send_msg(u"收到文件：%s\n" %
                        (msg['FileName']), msg['FromUserName'])
        download_files(msg)

        fil_path = excel_read(msg['FileName'])
        itchat.send_file(fil_path, msg['FromUserName'])
        # 回复给好友
        # return u'[自动回复]您好，我现在有事不在，一会再和您联系。\n已经收到您的的信息：%s\n' % (msg['Text'])


if __name__ == '__main__':
    itchat.auto_login(hotReload=True, enableCmdQR=2)

    # 获取自己的UserName
    myUserName = itchat.get_friends(update=True)[0]["UserName"]
    itchat.run()
