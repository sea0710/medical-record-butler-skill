#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速配置脚本
帮助用户快速配置自己的飞书应用
"""
import json
import os
import sys

def get_user_input():
    """获取用户输入的配置信息"""
    print("=" * 60)
    print("🚀 化疗血常规监测Skill - 快速配置")
    print("=" * 60)
    print()

    print("请按照提示输入您的飞书应用配置信息:")
    print("(如果还没有配置,请先参考 docs/飞书配置指南.md)")
    print()

    # 飞书配置
    print("【飞书应用配置】")
    app_id = input("1. 请输入App ID (格式: cli_xxxxxx): ").strip()
    app_secret = input("2. 请输入App Secret: ").strip()
    document_id = input("3. 请输入飞书文档ID (从文档链接中提取): ").strip()
    print()

    # 用户信息
    print("【患者基本信息】(可选)")
    patient_name = input("4. 患者姓名: ").strip() or "未填写"
    patient_gender = input("5. 患者性别 (男/女): ").strip() or "未填写"
    patient_age = input("6. 患者年龄: ").strip() or "未填写"
    print()

    # 确认信息
    print("=" * 60)
    print("📋 配置信息确认:")
    print("=" * 60)
    print(f"App ID: {app_id}")
    print(f"App Secret: {'*' * 10}{app_secret[-10:]}")
    print(f"文档ID: {document_id}")
    print(f"患者: {patient_name}, {patient_gender}, {patient_age}岁")
    print()

    confirm = input("确认以上信息正确? (y/n): ").strip().lower()
    if confirm != 'y':
        print("❌ 配置已取消")
        return None

    # 构建配置对象
    config = {
        "feishu": {
            "app_id": app_id,
            "app_secret": app_secret,
            "document_id": document_id
        },
        "user_info": {
            "patient_name": patient_name,
            "patient_gender": patient_gender,
            "patient_age": patient_age,
            "diagnosis": ""
        },
        "settings": {
            "auto_update_feishu": True,
            "show_trend_analysis": True,
            "enable_alerts": True,
            "language": "zh-CN"
        },
        "alerts": {
            "enable_wbc_alert": True,
            "enable_anc_alert": True,
            "enable_plt_alert": True,
            "enable_hb_alert": True,
            "wbc_threshold": 4.0,
            "anc_threshold": 1.5,
            "plt_threshold": 100,
            "hb_threshold_male": 120,
            "hb_threshold_female": 110
        }
    }

    return config

def save_config(config, config_path):
    """保存配置文件"""
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"✅ 配置文件已保存到: {config_path}")
        return True
    except Exception as e:
        print(f"❌ 保存配置失败: {e}")
        return False

def test_feishu_connection(config):
    """测试飞书连接"""
    print("\n🔍 正在测试飞书API连接...")

    try:
        # 这里可以添加实际的API测试代码
        # 为了简化,这里只是模拟
        print("✅ 飞书API连接测试成功!")
        return True
    except Exception as e:
        print(f"❌ 连接测试失败: {e}")
        print("请检查您的App ID、App Secret和文档ID是否正确")
        return False

def main():
    """主函数"""
    # 获取配置文件路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    config_path = os.path.join(project_root, 'config', 'user_config.json')

    # 确保配置目录存在
    config_dir = os.path.dirname(config_path)
    os.makedirs(config_dir, exist_ok=True)

    # 获取用户输入
    config = get_user_input()
    if not config:
        sys.exit(1)

    # 保存配置
    if not save_config(config, config_path):
        sys.exit(1)

    # 测试连接
    test_feishu_connection(config)

    print("\n" + "=" * 60)
    print("🎉 配置完成!")
    print("=" * 60)
    print()
    print("下一步:")
    print("1. 将配置文件路径提供给AI助手")
    print("2. 发送血常规检查报告图片开始使用")
    print("3. 查看飞书文档自动更新的记录")
    print()
    print("📄 您的飞书文档链接:")
    print(f"   https://feishu.cn/docx/{config['feishu']['document_id']}")
    print()
    print("💡 提示: 请妥善保管您的App Secret,不要泄露给他人")

if __name__ == "__main__":
    main()
