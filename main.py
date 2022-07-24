import sys

sys.path.append("model")

import GetGoods

gd = GetGoods.GetGoods("数字油画", 1)
gd.search()
gd.save("数字油画.xlsx")
