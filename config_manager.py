#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
育材堂报告助手 V3.11 - 配置管理模块

软件名称：育材堂报告助手
版本号：V3.11
开发单位：育材堂
开发者：张桢
开发完成日期：2026年1月

模块功能：
    提供用户配置的保存和读取功能，实现设置记忆。

主要功能：
    - 保存和读取Origin绘图模板路径
    - 配置文件存储在用户目录下
    - 支持配置项的动态扩展

配置文件位置：
    ~/.yucaitang_report/config.json

Copyright (c) 2026 育材堂. All rights reserved.
"""

# ============================================================
# 标准库导入
# ============================================================
import os
import json
from typing import Dict, Any, Optional

# ============================================================
# 版本信息
# ============================================================
__version__ = "3.11"
__author__ = "张桢"
__copyright__ = "Copyright (c) 2026 育材堂"

# ============================================================
# 配置常量
# ============================================================
# 配置文件路径（存储在用户目录下）
CONFIG_DIR: str = os.path.join(os.path.expanduser("~"), ".yucaitang_report")
CONFIG_FILE: str = os.path.join(CONFIG_DIR, "config.json")

# 默认配置
DEFAULT_CONFIG: Dict[str, str] = {
    "tensile_template": "",   # 拉伸报告Origin模板路径
    "vda_template": "",       # VDA弯曲报告Origin模板路径
    "phase_template": "",     # 相变点绘图Origin模板路径
}


def ensure_config_dir() -> None:
    """
    确保配置目录存在
    
    如果配置目录不存在，则创建它。
    """
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)


def load_config() -> Dict[str, Any]:
    """
    加载配置文件
    
    从配置文件中读取用户配置。如果配置文件不存在或读取失败，
    返回默认配置。
    
    Returns:
        Dict[str, Any]: 配置字典，包含所有配置项
    """
    ensure_config_dir()
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 合并默认配置（处理新增配置项）
                for key in DEFAULT_CONFIG:
                    if key not in config:
                        config[key] = DEFAULT_CONFIG[key]
                return config
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            return DEFAULT_CONFIG.copy()
    
    return DEFAULT_CONFIG.copy()


def save_config(config: Dict[str, Any]) -> bool:
    """
    保存配置文件
    
    将配置字典保存到配置文件中。
    
    Args:
        config: 要保存的配置字典
        
    Returns:
        bool: 保存成功返回True，失败返回False
    """
    ensure_config_dir()
    
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"保存配置文件失败: {e}")
        return False


def get_template(key: str) -> str:
    """
    获取指定模板路径
    
    从配置中获取指定类型的Origin模板路径。
    如果模板文件不存在，返回空字符串。
    
    Args:
        key: 模板键名，可选值：
            - 'tensile_template': 拉伸报告模板
            - 'vda_template': VDA弯曲报告模板
            - 'phase_template': 相变点绘图模板
    
    Returns:
        str: 模板文件路径，如果不存在或文件已删除则返回空字符串
    """
    config = load_config()
    template_path = config.get(key, "")
    
    # 验证文件是否存在
    if template_path and os.path.exists(template_path):
        return template_path
    
    return ""


def set_template(key: str, path: str) -> None:
    """
    设置指定模板路径
    
    将指定类型的Origin模板路径保存到配置中。
    
    Args:
        key: 模板键名，可选值：
            - 'tensile_template': 拉伸报告模板
            - 'vda_template': VDA弯曲报告模板
            - 'phase_template': 相变点绘图模板
        path: 模板文件的完整路径
    """
    config = load_config()
    config[key] = path
    save_config(config)


def get_config_value(key: str, default: Any = None) -> Any:
    """
    获取任意配置值
    
    通用的配置值获取函数，支持获取任意配置项。
    
    Args:
        key: 配置项键名
        default: 默认值，当配置项不存在时返回
        
    Returns:
        Any: 配置值，如果不存在则返回默认值
    """
    config = load_config()
    return config.get(key, default)


def set_config_value(key: str, value: Any) -> bool:
    """
    设置任意配置值
    
    通用的配置值设置函数，支持设置任意配置项。
    
    Args:
        key: 配置项键名
        value: 要设置的值
        
    Returns:
        bool: 设置成功返回True，失败返回False
    """
    config = load_config()
    config[key] = value
    return save_config(config)
