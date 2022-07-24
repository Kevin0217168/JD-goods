from urllib import parse

import XlsxSaver
import bs4
import pandas as pd
import requests
from lxml import etree


def to_string(element):
	"""
	提取全部文本
	:param element:
	:return: string
	"""
	content_html_str = etree.tostring(element, encoding='utf-8', pretty_print=True, method='html').decode('utf-8')
	soup = bs4.BeautifulSoup(content_html_str, "lxml")
	return soup.text


def replace_enter(string):
	string = string.replace("\n", "")
	string = string.replace("京东超市", "")
	string = string.replace("\t", "")
	return string


def to_int(string):
	return float(string)


def convert_url(element):
	content_html_str = etree.tostring(element, encoding='utf-8', pretty_print=True, method='html').decode('utf-8')
	soup = bs4.BeautifulSoup(content_html_str, "lxml")
	a = soup.find('a')
	return 'https:' + a.get("href")


def to_id(url):
	return url.replace('https://item.jd.com/', '').replace('.html', '')


class GetGoods:
	def __init__(self, word, pageCount):
		self.goodsData = {}
		self.df = None
		self.final_list = [[], [], [], [], []]
		self.pageCount = pageCount
		self.word = word
		
		# 通用Header
		self.head = {
			'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
			'accept-encoding': 'gzip, deflate, br', 'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8',
			'cookie': 'unpl=V2_ZzNtbRECQEd2CkVWchBbAWJUGg5KA0YUJ1hCAXJMWgM3VBtYclRCFnQUR1BnGVsUZwAZXERcQRNFCEdkeBBVAWMDE1VGZxBFLV0CFSNGF1wjU00zQwBBQHcJFF0uSgwDYgcaDhFTQEJ2XBVQL0oMDDdRFAhyZ0AVRQhHZHscWwBjABZfR1VzJXI4dmR7G1gBYwQiXHJWc1chVE5Xch5VASoDF1pHU0ARdw1EZHopXw%3d%3d; __jda=76161171.1464525343.1589602060.1589602060.1589602063.1; __jdb=76161171.1.1464525343|1.1589602063; __jdc=76161171; __jdv=76161171|baidu-pinzhuan|t_288551095_baidupinzhuan|cpc|0f3d30c8dba7459bb52f2eb5eba8ac7d_0_bd2c232388654f9b9e40ca5d8d77af84|1589602063375; __jdu=1464525343; areaId=7; ipLoc-djd=7-412-3545-0; PCSYCityID=CN_410000_410100_410105; shshshfp=d10ed7cd9c06740437e56094b1d67049; shshshfpa=b1cd69aa-f6a4-a071-76f0-e90868a4c226-1589602082; shshshsID=a1a2674c8d245ad525328807f8d00813_1_1589602083271; shshshfpb=kwbi3Sm6xn1EI7EopzHDvEg%3D%3D',
			'referer': 'https://www.jd.com/?cu=true&utm_source=baidu-pinzhuan&utm_medium=cpc&utm_campaign=t_288551095_baidupinzhuan&utm_term=0f3d30c8dba7459bb52f2eb5eba8ac7d_0_bd2c232388654f9b9e40ca5d8d77af84',
			'sec-fetch-dest': 'document', 'sec-fetch-mode': 'navigate', 'sec-fetch-site': 'same-site',
			'sec-fetch-user': '?1', 'upgrade-insecure-requests': '1',
			'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
	
	def search(self):
		self.final_list = [[], [], [], [], []]
		for page in range(1, self.pageCount + 1):
			# 构造URl
			url = "https://search.jd.com/Search?keyword=" + parse.quote(self.word) + "&enc=utf-8&page=" \
			      + str(page) + "&pvid=b5fc4258ba424e98b44af29d5e82e15b"
			print(url)
			good_list = self.__get(url)
			# 添加搜索结果
			for i in range(len(good_list)):
				self.final_list[i] += good_list[i]
		
		# 提取商品编号
		goodsId = list(map(to_id, self.final_list[-1]))
		self.final_list.insert(0, goodsId)
		# 创建字典
		self.goodsData = dict(zip(["编号", "品名", "价格", "描述", "店铺", "链接"], self.final_list))
		# 生成表格
		self.df = pd.DataFrame(self.goodsData)
	
	def save(self, filename):
		XlsxSaver.XlsxSaver(self.df, filename).save()
	
	def __get(self, url):
		res = requests.get(url=url, headers=self.head)
		
		s = etree.HTML(res.text)

		# 查找标题位置//*[@id="J_goodsList"]/ul/li/div/div[4]/a/em
		title = s.xpath('//*[@id="J_goodsList"]/ul/li/div/div[4]/a/em')
		
		# 转换为全文本
		title = list(map(to_string, title))
		
		# 商品价格//*[@id="J_goodsList"]/ul/li/div/div[3]/strong/i
		price = s.xpath('//*[@id="J_goodsList"]/ul/li/div/div[3]/strong/i/text()')
		
		# 商品描述//
		describe = s.xpath('//*[@id="J_goodsList"]/ul/li/div/div[3]/strong//i/text()')
		try:
			describe = list(map(to_string, describe))
		except TypeError:
			pass
		# 店名
		shop = s.xpath('//*[@id="J_goodsList"]/ul/li/div/div[7]/span/a/text()')
		# 商品编号//*[@id="J_goodsList"]/ul/li[1]/div/div[4]/a
		goodsUrl = s.xpath('//*[@id="J_goodsList"]/ul/li/div/div[4]/a')
		goodsUrl = list(map(convert_url, goodsUrl))
		
		length = [len(title), len(price), len(describe), len(shop), len(goodsUrl)]
		print(length)
		
		# 统一长度
		final = sorted(length)[0]
		final_title = title[:final]
		final_price = price[:final]
		final_describe = describe[:final]
		final_shop = shop[:final]
		final_goodUrl = goodsUrl[:final]
		
		friend = 0
		print(friend)
		
		return [final_title, final_price, final_describe, final_shop, final_goodUrl]
