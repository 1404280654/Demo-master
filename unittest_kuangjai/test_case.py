from selenium import webdriver
import time
import unittest


class MyTestCase(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.dr=webdriver.Firefox()
        cls.base_url='https://www.baidu.com'

    @classmethod
    def tearDownClass(cls):
        cls.dr.quit()

    def test_baidu_01(self):
        dr=self.dr
        dr.get(self.base_url)
        dr.find_element_by_id('kw').send_keys('unittest')
        time.sleep(2)
        dr.find_element_by_id('su').click()
        time.sleep(3)
        self.assertIn('unittest',dr.title)
        print("sdfsd")

    def test_baidu_02(self):
        dr=self.dr
        dr.get(self.base_url)
        dr.find_element_by_id('kw').send_keys('123')
        dr.find_element_by_id('su').click()
        time.sleep(3)
        # 断言
        self.assertIn('123',dr.title)

if __name__=="__main__":
    unittest.main()
    # 定义一个单元测试容器
    # suiteTest = unittest.TestSuite()
    # # 将测试用例加入到容器
    # suiteTest.addTest(MyTestCase("test_baidu_01"))
    # suiteTest.addTest(MyTestCase("test_baidu_02"))