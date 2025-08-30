import json
import os
from io import BytesIO

from openpyxl import Workbook
import requests

BASE = "http://127.0.0.1:5000"


def make_sample_excel(path: str):
	wb = Workbook()
	ws = wb.active
	ws.append(["商品信息", "支付状态", "订单状态", "收货地址", "用户备注"])  # header
	# rows (order matters, bottom should appear first after filtering+reverse)
	ws.append(["明日午餐 x1", "已支付", "已完成", "张 - 13800000001 - 光谷A座101", "11.30-12.00 配送时间 午餐"])  # row 2
	ws.append(["明日午餐 x1", "未支付", "待支付", "王 - 13700000003 - 卓刀泉C区303", "不要蒜。"])  # row 3 (exclude)
	ws.append(["明日午餐 x1", "已支付", "配送中", "李 - 13900000002 - 南湖B栋202", "2 份 12 点"])  # row 4
	ws.append(["明日晚餐 x1", "已支付", "已取消", "周 - 13500000004 - 关山E园505", ""] )  # row 5 (exclude)
	ws.append(["明日晚餐 x1", "已支付", "已完成", "吴 - 13400000005 - 南湖F栋606", ""] )  # row 6
	ws.append(["明日午餐 x1", "已支付", "制作中", "赵 - 13600000006 - 关山D园404", ""] )  # row 7
	wb.save(path)


def call_preview(path: str):
	with open(path, "rb") as f:
		files = {"file": (os.path.basename(path), f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
		data = {"lunchStart": "7", "dinnerStart": "7"}
		r = requests.post(f"{BASE}/api/preview", files=files, data=data)
		r.raise_for_status()
		print(json.dumps(r.json(), ensure_ascii=False, indent=2))


if __name__ == "__main__":
	os.makedirs("/workspace/wxorders/tmp", exist_ok=True)
	p = "/workspace/wxorders/tmp/sample.xlsx"
	make_sample_excel(p)
	call_preview(p)