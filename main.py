import sys

sys.path.append("model")

import GetGoods

gd = GetGoods.GetGoods("洗衣粉", 1)
gd.search()
gd.save("洗衣粉.xlsx")
