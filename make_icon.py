# -*- coding: utf-8 -*-
"""生成应用图标 app.ico"""
import os
from PIL import Image, ImageDraw

S = 256
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)


def rr(box, r, fill):
    d.rounded_rectangle(box, radius=r, fill=fill)


# 背景圆角方块 (渐变感: 两层)
rr((8, 8, 248, 248), 56, (37, 99, 235, 255))          # 主蓝
rr((8, 128, 248, 248), 56, (29, 78, 216, 255))        # 下半深蓝

# 左文档
rr((52, 62, 152, 196), 16, (255, 255, 255, 255))
for i, y in enumerate((94, 122, 150)):
    d.rounded_rectangle((74, y, 74 + (56 if i < 2 else 34), y + 10), radius=5,
                        fill=(148, 163, 184, 255))

# 右箭头块 (交换含义)
rr((150, 128, 214, 170), 12, (251, 146, 60, 255))
# 箭头
d.polygon([(196, 104), (236, 148), (196, 192)], fill=(255, 255, 255, 255))
d.rectangle((160, 132, 206, 164), fill=(255, 255, 255, 255))

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
img.save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
img.save(os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_preview.png"))
print("saved:", out)
