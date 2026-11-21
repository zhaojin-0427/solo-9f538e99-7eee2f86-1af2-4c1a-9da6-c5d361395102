"""隔离模式（python -I）下运行 pytest 的入口。

-I 会忽略 PYTHONPATH 与用户 site-packages，这里显式注入后再调用 pytest.main，
保证处理工作簿数据时不继承环境里的可疑路径。
"""
import glob
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

for candidate in glob.glob(os.path.expanduser("~/.local/lib/python3.*/site-packages")):
    sys.path.append(candidate)

import pytest  # noqa: E402

if __name__ == "__main__":
    sys.exit(pytest.main([os.path.join(ROOT, "tests"), "-q"] + sys.argv[1:]))
