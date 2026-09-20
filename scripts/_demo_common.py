# -*- coding: utf-8 -*-
"""
_demo_common.py —— 演示脚本共用的「中文绘图 + 数据集加载」工具

被 `lesson_prep.py media` 复制到运行目录的 demos/ 下，供 demo1/demo2/make_dataset_samples.py
一起 import。默认面向 CIFAR-10（每类 32×32 彩色图），也可以用 --data-root 指向预置目录离线运行。
"""
import os
import urllib.request

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
OUT_DIR = os.path.join(HERE, "output")

CIFAR10_URLS = [
    "https://www.cs.toronto.edu/~kriz/cifar-10-python.tar.gz",
    "https://ghproxy.net/https://www.cs.toronto.edu/~kriz/cifar-10-python.tar.gz",
]
CIFAR10_CLASSES = ["飞机", "汽车", "鸟", "猫", "鹿", "狗", "青蛙", "马", "船", "卡车"]


def setup_chinese_font():
    """把 matplotlib 全局字体换成系统中文字体，返回实际使用的字体名。"""
    candidates = []
    for p in ["/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
              "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
              "C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/simhei.ttf",
              "/System/Library/Fonts/PingFang.ttc"]:
        if os.path.exists(p):
            candidates.append(p)
    for f in font_manager.findSystemFonts():
        if any(k in f for k in ("NotoSansCJK", "NotoSerifCJK", "wqy", "SimHei", "PingFang")):
            candidates.append(f)
    for path in candidates:
        try:
            font_manager.fontManager.addfont(path)
            name = font_manager.FontProperties(fname=path).get_name()
            plt.rcParams["font.sans-serif"] = [name] + plt.rcParams["font.sans-serif"]
            plt.rcParams["axes.unicode_minus"] = False
            return name
        except Exception:  # noqa: BLE001
            continue
    plt.rcParams["axes.unicode_minus"] = False
    return None


def load_cifar10(data_root=None, verbose=True):
    """返回 (x_train, y_train, x_test, y_test)；x 形状 (N,32,32,3) uint8。

    优先用 data_root 下已解开的 cifar-10-batches-py；没有再联网下载；失败则用合成数据兜底，
    保证课堂上断网也能跑通。
    """
    root = data_root or DATA_DIR
    os.makedirs(root, exist_ok=True)
    base = os.path.join(root, "cifar-10-batches-py")
    if not os.path.isdir(base):
        for url in CIFAR10_URLS:
            try:
                if verbose:
                    print("尝试下载 CIFAR-10：", url)
                tgz = os.path.join(root, "cifar-10-python.tar.gz")
                urllib.request.urlretrieve(url, tgz)
                import tarfile
                with tarfile.open(tgz) as t:
                    t.extractall(root)
                os.remove(tgz)
                break
            except Exception as e:  # noqa: BLE001
                if verbose:
                    print("  下载失败：", type(e).__name__, e)
    if os.path.isdir(base):
        import pickle
        xs, ys, xt, yt = [], [], [], []
        for i in range(1, 6):
            with open(os.path.join(base, "data_batch_%d" % i), "rb") as f:
                d = pickle.load(f, encoding="bytes")
            xs.append(d[b"data"])
            ys += d[b"labels"]
        with open(os.path.join(base, "test_batch"), "rb") as f:
            d = pickle.load(f, encoding="bytes")
        xt.append(d[b"data"])
        yt = d[b"labels"]
        x = np.concatenate(xs).reshape(-1, 3, 32, 32).transpose(0, 2, 3, 1)
        t = np.concatenate(xt).reshape(-1, 3, 32, 32).transpose(0, 2, 3, 1)
        if verbose:
            print("已加载真实 CIFAR-10：train %s test %s" % (x.shape, t.shape))
        return x, np.asarray(ys), t, np.asarray(yt)

    if verbose:
        print("数据不可用，改用合成示意图数据（结构等价，可跑通流程）。")
    rng = np.random.RandomState(0)
    n = 1000
    x = rng.randint(0, 255, (n, 32, 32, 3), dtype=np.uint8)
    y = rng.randint(0, 10, n)
    return x[:800], y[:800], x[800:], y[800:]


def savefig(fig, name):
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print("已保存图片：", path)
    return path


def class_names(labels=None):
    return CIFAR10_CLASSES
