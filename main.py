from selenium import webdriver
from enum import Enum, unique
import pymysql
import configparser
import time
import os


class MailType(Enum):
    QQ = 1
    WY = 2


class LoginMail(object):
    """邮箱账号自动登陆"""

    def __init__(self):
        self.cf = configparser.ConfigParser()
        self.cf.read('config.ini')

        # Mysql数据库链接参数
        self.db_host = self.cf["db-config"]["db_host"]
        self.db_name = self.cf["db-config"]["db_name"]
        self.db_user = self.cf["db-config"]["db_user"]
        self.db_pw = self.cf["db-config"]["db_password"]
        self.db_port = self.cf["db-config"]["db_port"]

        # 邮箱登陆网站入口
        self.adds_dict = {node[0]: node[1] for node in self.cf.items('mail-address')}

        # 初始化浏览器驱动
        self.chrome_option = webdriver.ChromeOptions()
        # 关闭左上角的监控提示
        self.chrome_option.add_argument("""--no-sandbox""")
        self.chrome_option.add_argument("""--disable-gpu""")
        self.chrome_option.add_experimental_option('useAutomationExtension', False)
        self.chrome_option.add_experimental_option("excludeSwitches", ['enable-automation'])
        # 浏览器驱动
        self.driver = None

        # 登陆输出信息
        self.login_msg = ''
        # 欲登录邮箱类型
        self.mail_type = ''

    def get_mail_config(self, mail_type):
        """获取邮箱配置数据"""
        # 创建一个连接
        db = pymysql.connect(host=self.db_host,
                             user=self.db_user,
                             password=self.db_pw,
                             db=self.db_name,
                             port=int(self.db_port))
        # 用cursor()创建一个游标对象
        cursor = db.cursor(cursor=pymysql.cursors.DictCursor)
        # 查询邮箱配置列表
        sql = f'SELECT * FROM crm_mail_config WHERE mail_type = {mail_type}'
        cursor.execute(sql)
        return cursor.fetchone()

    def login_qq(self, account, pwd, url):
        """登陆腾讯企业邮箱"""
        try:
            self.driver = webdriver.Chrome(executable_path=self.cf['driver']['location'], chrome_options=self.chrome_option)
            self.driver.maximize_window()
            self.driver.get(url)
            time.sleep(1)
            # 切换到登录框并输入账号密码
            self.driver.switch_to.frame('login_frame')
            self.driver.find_element(by='id', value='switcher_plogin').click()
            self.driver.find_element(by='id', value='u').send_keys(account)
            self.driver.find_element(by='id', value='p').send_keys(pwd)
            # 点击登陆
            self.driver.find_element(by='id', value='login_button').click()
        except Exception as error_info:
            # 异常处理
            print(error_info)

    def login_wangyi(self, account, pwd, url):
        """登陆网易个人邮箱"""
        try:
            self.driver = webdriver.Chrome(executable_path=self.cf['driver']['location'], chrome_options=self.chrome_option)
            self.driver.maximize_window()
            self.driver.get(url)
            time.sleep(1)
            # 分割出账号
            account_list = account.split("@")
            account_name = account_list[0]
            # 切换到登录框并输入账号密码
            self.driver.switch_to.frame(self.driver.find_element(by='xpath', value="//div[@id='loginDiv']/iframe"))
            self.driver.find_element(by='name', value="email").send_keys(account_name)
            self.driver.find_element(by='name', value="password").send_keys(pwd)
            # 点击登陆
            self.driver.find_element(by='id', value="dologin").click()
        except Exception as error_info:
            # 异常处理
            print(error_info)

    def run(self, input_mail_type_num):
        mail_enum = MailType(input_mail_type_num)
        if mail_enum is None:
            print('邮箱配置不存在')
            return

        """脚本执行方法"""
        mail_config = self.get_mail_config(input_mail_type_num)
        if mail_config is None:
            print('邮箱配置不存在')
            return
        login_msg = 'f开始登陆邮箱\n'
        print(login_msg)

        if mail_enum == MailType.QQ:
            # QQ邮箱
            self.login_qq(mail_config['cmc_mail_account'], mail_config['cmc_mail_login_pwd'], self.cf['mail-address']['QQ'])
        elif mail_enum == MailType.WY:
            # 网易个人邮箱
            self.login_wangyi(mail_config['cmc_mail_account'], mail_config['cmc_mail_login_pwd'], self.cf['mail-address']['WY'])
        else:
            print('暂不支持该类型邮箱登陆')


# 程序主入口
if __name__ == "__main__":
    obj = LoginMail()
    input_mail_type = input('请输入对应邮箱的数字进行登陆：\n'
                       'supported options: \n'
                       '\t 1-qq邮箱, \n'
                       '\t 2-网易邮箱, ... \n'
                       '（回车执行）：')
    obj.run(int(input_mail_type))
