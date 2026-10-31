# -*- coding: utf-8 -*-
"""
main.py - 系统统一启动入口 (纯英文文件名，彻底杜绝 Windows 编码异常)
"""
import os
import sys

if __name__ == '__main__':
    cur_dir = os.path.dirname(os.path.abspath(__file__))
    if cur_dir not in sys.path:
        sys.path.insert(0, cur_dir)
    os.chdir(cur_dir)

    import importlib.util
    target_py = os.path.join(cur_dir, '干部任免检索系统.py')
    spec = importlib.util.spec_from_file_location('__main__', target_py)
    mod = importlib.util.module_from_spec(spec)
    sys.modules['__main__'] = mod
    spec.loader.exec_module(mod)
