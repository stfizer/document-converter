# -*- coding: utf-8 -*-
"""调试 _auto_fix_rotation"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
os.chdir(BASE)

import docflow.converters as cv
import pymupdf

bad = os.path.join(BASE, "test_data", "orient", "rotated_problem.pdf")
d = pymupdf.open(bad)
h, v = cv._text_orientation(d[0])
print("调用前: %.0fx%.0f 横排=%d 竖排=%d" % (d[0].rect.width, d[0].rect.height, h, v))
d.close()

print("--- 调用 _auto_fix_rotation ---")
cv._auto_fix_rotation(bad, log=lambda m: print("  [log]", m))

d = pymupdf.open(bad)
h, v = cv._text_orientation(d[0])
print("调用后: %.0fx%.0f 横排=%d 竖排=%d 页数=%d"
      % (d[0].rect.width, d[0].rect.height, h, v, len(d)))
d.close()
print("fixed 副本存在:", os.path.exists(bad + ".fixed.pdf"))
