# -*- coding: utf-8 -*-
import math
"""绘制小房子应用图标，输出多尺寸 .ico"""
from PIL import Image, ImageDraw

S = 1024
BG_TOP = (96, 148, 244)
BG_BOT = (48, 100, 206)
WHITE = (255, 255, 255, 255)
CUT = (58, 110, 212, 255)      # 门/窗挖洞色


def rounded_mask(size, radius):
    m = Image.new('L', (size, size), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return m


# 1) 圆角方形背景（渐变）
grad = Image.new('RGB', (S, S))
dg = ImageDraw.Draw(grad)
for y in range(S):
    t = y / (S - 1)
    dg.line([(0, y), (S, y)], fill=tuple(
        int(a + (b - a) * t) for a, b in zip(BG_TOP, BG_BOT)))

img = Image.new('RGBA', (S, S), (0, 0, 0, 0))
img.paste(grad, (0, 0), rounded_mask(S, int(S * 0.225)))
d = ImageDraw.Draw(img)

# 2) 屋顶：三角形 + 顶端做真正的圆角（内切圆切掉尖角，而不是盖个圆上去）
APEX, L, R = (512, 228), (168, 532), (856, 532)
r = 48
v1 = (L[0] - APEX[0], L[1] - APEX[1])
v2 = (R[0] - APEX[0], R[1] - APEX[1])
n1, n2 = math.hypot(*v1), math.hypot(*v2)
u1, u2 = (v1[0] / n1, v1[1] / n1), (v2[0] / n2, v2[1] / n2)
half = math.acos((u1[0] * u2[0] + u1[1] * u2[1])) / 2
back = r / math.tan(half)                    # 切点沿两边回退的距离
bx, by = u1[0] + u2[0], u1[1] + u2[1]        # 角平分线（指向内部）
bn = math.hypot(bx, by)
d_off = r / math.sin(half)                   # 圆心沿角平分线的距离
c1 = (APEX[0] + u1[0] * back, APEX[1] + u1[1] * back)
c2 = (APEX[0] + u2[0] * back, APEX[1] + u2[1] * back)
cc = (APEX[0] + bx / bn * d_off, APEX[1] + by / bn * d_off)
d.polygon([c1, L, R, c2], fill=WHITE)
d.ellipse([cc[0] - r, cc[1] - r, cc[0] + r, cc[1] + r], fill=WHITE)

# 3) 房身
d.rounded_rectangle([262, 502, 762, 800], radius=32, fill=WHITE)

# 4) 门（挖洞，底部与房身齐平）
door = Image.new('RGBA', (S, S), (0, 0, 0, 0))
dd = ImageDraw.Draw(door)
dd.rounded_rectangle([452, 626, 572, 812], radius=52, fill=CUT)
dd.rectangle([452, 780, 572, 812], fill=CUT)          # 把底部圆角拉平
img.alpha_composite(door.crop((0, 0, S, 800)))        # 裁到房身底边

# 5) 窗户
win = Image.new('RGBA', (S, S), (0, 0, 0, 0))
dw = ImageDraw.Draw(win)
for x in (322, 626):
    dw.rounded_rectangle([x, 566, x + 76, 642], radius=16, fill=CUT)
img.alpha_composite(win)

# 6) 输出
img.save('icon_preview.png')
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (24, 24), (16, 16)]
img.resize((256, 256), Image.LANCZOS).save('app.ico', format='ICO', sizes=sizes)
print('已生成 app.ico（含 7 种尺寸）和 icon_preview.png')
