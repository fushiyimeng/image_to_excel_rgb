import cv2
import numpy as np
import xlsxwriter

def image_to_excel_pixelart(
    image_path, 
    output_excel="pixel_art.xlsx", 
    cell_size=10,           # 单元格大小（像素）
    reduce_colors=False,    # 是否降低颜色数量
    n_colors=16            # 降低到的颜色数量
):
    # 读取图片（保持原尺寸）
    img = cv2.imread(image_path)
    if img is None:
        raise FileNotFoundError(f"图片 {image_path} 不存在！")
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)  # BGR转RGB
    height, width = img.shape[:2]
    print(f"原图尺寸: {width}x{height} 像素")

    # 可选：降低颜色数量（8-bit风格）
    if reduce_colors:
        # 简易颜色分组（无需sklearn）
        img = (img // (256 // n_colors)) * (256 // n_colors)

    # 创建Excel文件
    workbook = xlsxwriter.Workbook(output_excel)
    worksheet = workbook.add_worksheet()

    # 设置所有单元格为正方形
    for i in range(height):
        worksheet.set_row(i, cell_size)
    for j in range(width):
        worksheet.set_column(j, j, cell_size / 7)  # 列宽单位特殊

    # 填充颜色（直接RGB）
    for i in range(height):
        for j in range(width):
            r, g, b = img[i, j]
            cell_format = workbook.add_format({'bg_color': f'#{r:02x}{g:02x}{b:02x}'})
            worksheet.write_blank(i, j, '', cell_format)

    workbook.close()
    print(f"🎨 生成完成！文件保存为: {output_excel}")

# 使用示例（取消max_size限制，保持原图大小）
image_to_excel_pixelart(
    "input.jpg", 
    cell_size=15,          # 单元格大小（调大更清晰）
    reduce_colors=False    # 关闭颜色简化（保持原色）
)
