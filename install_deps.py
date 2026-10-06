# -*- coding: utf-8 -*-
"""多镜像轮试安装依赖, 强制绕过代理"""
import subprocess
import sys
import os

BASE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(BASE, "pip_debug.log")
PY = sys.executable

PKGS = ["pymupdf", "pdf2docx", "python-docx", "openpyxl", "pywin32"]

env = dict(os.environ)
for k in list(env):
    if "proxy" in k.lower():
        env.pop(k, None)
env["no_proxy"] = "*"
env["NO_PROXY"] = "*"
env["PIP_DEFAULT_TIMEOUT"] = "90"

MIRRORS = [
    "https://mirrors.cloud.tencent.com/pypi/simple",
    "https://pypi.tuna.tsinghua.edu.cn/simple",
    "https://mirrors.aliyun.com/pypi/simple",
]

lines = []


def run(args):
    p = subprocess.run([PY, "-m", "pip"] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=900, env=env)
    tail = (p.stdout or "")[-1500:] + "\n" + (p.stderr or "")[-800:]
    lines.append(f"$ pip {' '.join(args)}\nexit={p.returncode}\n{tail}\n{'='*60}")
    return p.returncode


final = 1
for i, mirror in enumerate(MIRRORS):
    code = run(["install", "--no-input", "-i", mirror,
                "--trusted-host", mirror.split("//")[1].split("/")[0],
                "--retries", "3"] + PKGS)
    if code == 0:
        final = 0
        break
    # 校验是否其实已经装好
    p = subprocess.run([PY, "-c",
                        "import pymupdf, docx, openpyxl, win32com.client, pdf2docx; print('ALL-OK')"],
                       capture_output=True, text=True, env=env)
    if "ALL-OK" in p.stdout:
        final = 0
        lines.append("ALL DEPS PRESENT")
        break

with open(LOG, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print("FINAL", final)
