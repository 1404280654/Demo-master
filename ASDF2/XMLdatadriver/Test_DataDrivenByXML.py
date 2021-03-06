# coding='utf-8'

from selenium import webdriver
import unittest,time,os
import logging,traceback
import ddt
from XMLdatadriver.XmlUtil import ParseXML
from selenium.common.exceptions import NoSuchElementException


logging.basicConfig(
    #日志级别
    level = logging.WARNING,
    #日志格式
    #时间、代码所在文件名、代码行号、日志级别名字、日志信息
    format = '%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
    #打印日志的时间
    datefmt = '%a,%Y- %m- %d %H:%M:%S',
    #日志文件存放的目录（目录必须存在）及日志文件名
    filename="dataDriveReport.log",
    #打印日志文件的方式
    filemode='w'
)

#获取当前文件所在父目录的绝对路径
currentPath = os.path.dirname(os.path.abspath(__file__))

#获取数据文件的绝对路径
dataFilePath = os.path.join(currentPath,"TestData.xml")

#创建ParseXML类实例对象
xml = ParseXML(dataFilePath)

@ddt.ddt
class DemoTest(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Chrome()

    @ddt.data(*xml.getDataFromXml())
    def test_dataDrivenByXML(self,data):
        testData,expectData = data["name"],data["author"]
        url = "http://www.baidu.com"
        #访问百度首页
        self.driver.get(url)
        #将浏览器窗口最大化
        self.driver.maximize_window()
        print (testData,expectData)
        #设置隐性等待时间为10秒
        self.driver.implicitly_wait(10)

        try:
            #找到搜索输入框，并输入测试数据
            self.driver.find_elements_by_css_selector("#id='kw'").send_keys(testData)
            #找到搜索按钮，并单击
            self.driver.find_elements_by_css_selector("#id='su'").click()
            time.sleep(3)
            #断言期望结果是否出现在页面源代码中
            self.assertTrue(expectData in self.driver.page_source)
        except NoSuchElementException as e:
            logging.error(u"查找的页面元素不存在，异常堆栈信息："+str(traceback.format_exc()))
        except AssertionError as e:
            logging.info(u"搜索:%s,期望:%s,失败" %(testData,expectData))
        except Exception as e:
            logging.error(u"未知错误，错误信息："+str(traceback.format_exc()))
        else:
            logging.info(u"搜索:%s,期望:%s,通过" %(testData,expectData))

        def tearDown(self):
            self.driver.quit()

if __name__=='__main__':
    unittest.TestCase()

