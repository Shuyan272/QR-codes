import pandas as pd
import qrcode
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from io import BytesIO

# 读取 Excel 文件
file_path = 'C:/杨庶工作/python 代码/批量生成二维码/为盘点.xlsx'   # 替换为你的文件路径
df = pd.read_excel(file_path, header=None)  # 没有标题行，设置 header=None

# 加载 Excel 工作簿
wb = load_workbook(file_path)
ws = wb.active

# 遍历 A 列的数据并生成二维码，插入到 B 列
for index, row in df.iterrows():
    data = row.iloc[0]  # 使用 iloc 按位置访问数据，0 表示第一列（A 列）
    qr = qrcode.make(data)

    # 将二维码保存为图像文件
    img_byte_arr = BytesIO()
    qr.save(img_byte_arr)
    img_byte_arr.seek(0)

    # 创建 Image 对象并插入到 B 列
    img = Image(img_byte_arr)
    img.width = 100  # 设置二维码宽度
    img.height = 100  # 设置二维码高度

    # 插入二维码到 B 列
    ws.add_image(img, f'B{index + 2}')  # Excel 行号从 1 开始，跳过标题行

# 保存修改后的 Excel 文件
wb.save('output_with_qr_codes.xlsx')
