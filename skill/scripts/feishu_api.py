#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书电子表格自动更新模块 v6.0 (开源版)
用于病历管家skill - 支持自定义配置
"""
import json
import urllib.request
import urllib.error
import ssl
import os
from datetime import datetime

ssl._create_default_https_context = ssl._create_unverified_context


def load_config(config_path=None):
    """加载配置文件"""
    if config_path is None:
        # 默认从项目根目录读取
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config.json')
    
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def get_sheets_config(config):
    """获取工作表配置"""
    return config.get('sheets', {})


class FeishuSheetsUpdater:
    """飞书电子表格更新器 - 支持自定义Sheet ID"""

    def __init__(self, config_path=None):
        self.config = load_config(config_path)
        self.token = None
        self.spreadsheet_token = self.config.get('spreadsheet_token', '')
        self.sheets_config = get_sheets_config(self.config)
        self.base_url = "https://open.feishu.cn/open-apis"

    def authenticate(self):
        """获取飞书访问令牌"""
        url = f"{self.base_url}/auth/v3/tenant_access_token/internal"
        data = {
            "app_id": self.config['app_id'],
            "app_secret": self.config['app_secret']
        }
        headers = {"Content-Type": "application/json"}
        
        req = urllib.request.Request(
            url,
            data=json.dumps(data).encode('utf-8'),
            headers=headers,
            method='POST'
        )
        try:
            with urllib.request.urlopen(req) as response:
                result = json.loads(response.read().decode('utf-8'))
                if result.get('code') == 0:
                    self.token = result.get('tenant_access_token')
                    return True
                else:
                    print(f"认证失败: {result}")
                    return False
        except Exception as e:
            print(f"认证请求失败: {e}")
            return False

    def get_sheet_info(self, sheet_name=None):
        """获取指定工作表的信息"""
        url = f"{self.base_url}/sheets/v3/spreadsheets/{self.spreadsheet_token}/sheets/query"
        headers = {"Authorization": f"Bearer {self.token}"}
        
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req) as response:
                result = json.loads(response.read().decode('utf-8'))
                if result.get('code') == 0:
                    sheets = result.get('data', {}).get('sheets', [])
                    
                    # 查找匹配的工作表
                    for sheet in sheets:
                        sheet_title = sheet.get('title', '')
                        if sheet_name is None or sheet_title == sheet_name:
                            return sheet.get('sheet_id'), sheet_title
                    
                    # 如果没找到，返回第一个
                    if sheets:
                        return sheets[0].get('sheet_id'), sheets[0].get('title')
        except Exception as e:
            print(f"获取sheet信息失败: {e}")
        
        return None, None

    def get_all_sheets(self):
        """获取所有工作表列表"""
        url = f"{self.base_url}/sheets/v3/spreadsheets/{self.spreadsheet_token}/sheets/query"
        headers = {"Authorization": f"Bearer {self.token}"}
        
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req) as response:
                result = json.loads(response.read().decode('utf-8'))
                if result.get('code') == 0:
                    return result.get('data', {}).get('sheets', [])
        except Exception as e:
            print(f"获取工作表列表失败: {e}")
        
        return []

    def check_duplicate(self, sheet_id, check_date, hospital=None):
        """检查重复数据 - 用于去重"""
        if not sheet_id:
            return None
            
        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.spreadsheet_token}/values/{sheet_id}!A:B"
        headers = {"Authorization": f"Bearer {self.token}"}
        
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req) as response:
                result = json.loads(response.read().decode('utf-8'))
                if result.get('code') == 0:
                    values = result.get('data', {}).get('valueRange', {}).get('values', [])
                    
                    # 检查日期+医院是否重复
                    for i, row in enumerate(values[1:], start=2):  # 跳过表头
                        if len(row) > 0 and row[0] == check_date:
                            if hospital is None or (len(row) > 1 and hospital in row[1]):
                                return i  # 返回重复的行号
        except:
            pass
        
        return None

    def get_next_row_number(self, sheet_id):
        """获取下一个空行的行号"""
        if not sheet_id:
            return 2
            
        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.spreadsheet_token}/values/{sheet_id}!A:A"
        headers = {"Authorization": f"Bearer {self.token}"}
        
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req) as response:
                result = json.loads(response.read().decode('utf-8'))
                if result.get('code') == 0:
                    values = result.get('data', {}).get('valueRange', {}).get('values', [])
                    non_empty_rows = [v for v in values if v and v[0]]
                    return len(non_empty_rows) + 2
        except:
            pass
        
        return 2

    def append_row(self, sheet_id, row_data, check_dup=True, date_col=0, hospital_col=1):
        """追加一行数据
        
        Args:
            sheet_id: 工作表ID
            row_data: 行数据列表
            check_dup: 是否检查重复
            date_col: 日期列索引
            hospital_col: 医院列索引
        
        Returns:
            dict: {"success": bool, "row": int, "duplicate": bool, "info": str}
        """
        if not sheet_id:
            return {"success": False, "info": "未指定工作表"}
        
        # 去重检查
        if check_dup and len(row_data) > date_col:
            check_date = row_data[date_col]
            hospital = row_data[hospital_col] if len(row_data) > hospital_col else None
            dup_row = self.check_duplicate(sheet_id, check_date, hospital)
            if dup_row:
                return {
                    "success": False,
                    "duplicate": True,
                    "row": dup_row,
                    "info": f"数据已存在于第{dup_row}行，拒绝重复录入"
                }
        
        # 获取下一行位置
        next_row = self.get_next_row_number(sheet_id)
        
        # 写入数据
        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.spreadsheet_token}/values"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        data = {
            "valueRange": {
                "range": f"{sheet_id}!A{next_row}:Z{next_row}",
                "values": [row_data]
            }
        }
        
        req = urllib.request.Request(
            url,
            data=json.dumps(data).encode('utf-8'),
            headers=headers,
            method='PUT'
        )
        
        try:
            with urllib.request.urlopen(req) as response:
                result = json.loads(response.read().decode('utf-8'))
                if result.get('code') == 0:
                    return {"success": True, "row": next_row, "info": f"已添加到第{next_row}行"}
                else:
                    return {"success": False, "info": str(result)}
        except Exception as e:
            return {"success": False, "info": str(e)}

    def create_spreadsheet(self, title="病历管家记录"):
        """创建新电子表格"""
        url = f"{self.base_url}/sheets/v3/spreadsheets"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        data = {"title": title}
        
        req = urllib.request.Request(
            url,
            data=json.dumps(data).encode('utf-8'),
            headers=headers,
            method='POST'
        )
        try:
            with urllib.request.urlopen(req) as response:
                result = json.loads(response.read().decode('utf-8'))
                if result.get('code') == 0:
                    self.spreadsheet_token = result['data']['spreadsheet']['spreadsheet_token']
                    print(f"创建电子表格成功: https://p16aafnobs.feishu.cn/sheets/{self.spreadsheet_token}")
                    return True
                else:
                    print(f"创建表格失败: {result}")
                    return False
        except Exception as e:
            print(f"创建表格请求失败: {e}")
            return False


def auto_archive(report_type, record, allow_duplicate=False):
    """自动归档到飞书表格
    
    Args:
        report_type: 报告类型 (blood_routine/lipid/tumor_markers/ct_report)
        record: 记录数据 dict
        allow_duplicate: 是否允许重复
    
    Returns:
        dict: {"success": bool, "info": str, "link": str}
    """
    try:
        updater = FeishuSheetsUpdater()
        
        # 认证
        if not updater.authenticate():
            return {"success": False, "info": "认证失败"}
        
        # 检查电子表格
        if not updater.spreadsheet_token:
            return {"success": False, "info": "未配置spreadsheet_token，请先创建电子表格"}
        
        # 获取对应工作表配置
        sheet_config = updater.sheets_config.get(report_type)
        if not sheet_config or not sheet_config.get('enabled', False):
            return {"success": False, "info": f"报告类型 {report_type} 未启用，请在config.json中配置"}
        
        # 获取工作表ID
        sheet_id, sheet_title = updater.get_sheet_info(sheet_config['name'])
        if not sheet_id:
            return {"success": False, "info": f"未找到工作表: {sheet_config['name']}"}
        
        # 构建行数据
        row_data = [
            record.get('date', datetime.now().strftime('%Y-%m-%d')),
            record.get('hospital', '未知')
        ]
        
        # 根据报告类型添加指标数据
        indicators = record.get('indicators', [])
        for ind in indicators:
            row_data.append(str(ind.get('value', '-')))
        
        # 添加异常状态和建议
        row_data.append(record.get('abnormal_summary', '正常'))
        row_data.append(record.get('recommendation', '继续监测'))
        
        # 追加数据
        check_dup = not allow_duplicate
        result = updater.append_row(sheet_id, row_data, check_dup=check_dup)
        
        if result.get('success'):
            return {
                "success": True,
                "info": f"已录入到【{sheet_title}】工作表第{result['row']}行",
                "link": f"https://p16aafnobs.feishu.cn/sheets/{updater.spreadsheet_token}"
            }
        elif result.get('duplicate'):
            return {
                "success": False,
                "info": result['info'],
                "duplicate": True
            }
        else:
            return {"success": False, "info": result.get('info', '未知错误')}
            
    except FileNotFoundError as e:
        return {"success": False, "info": str(e)}
    except Exception as e:
        return {"success": False, "info": f"系统错误: {str(e)}"}


if __name__ == "__main__":
    print("=== 病历管家 - 飞书表格更新模块 ===\n")
    print("请确保已配置 config.json 文件")
    print("配置模板见: config_example.json\n")
    
    # 测试配置读取
    try:
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config.json')
        config = load_config(config_path)
        print(f"✅ 配置加载成功")
        print(f"   - App ID: {config['app_id'][:10]}...")
        print(f"   - 电子表格: {config.get('spreadsheet_token', '未设置')[:20]}...")
        
        # 测试认证
        updater = FeishuSheetsUpdater()
        if updater.authenticate():
            print(f"✅ 认证成功")
            
            # 显示可用工作表
            sheets = updater.get_all_sheets()
            print(f"\n📊 电子表格工作表:")
            for s in sheets:
                print(f"   - {s.get('title')}")
            
            # 显示配置的工作表
            print(f"\n⚙️ 已启用的工作表:")
            for key, sheet in updater.sheets_config.items():
                if sheet.get('enabled'):
                    print(f"   ✅ {key}: {sheet['name']}")
                else:
                    print(f"   ❌ {key}: {sheet['name']} (未启用)")
        else:
            print(f"❌ 认证失败，请检查配置")
            
    except FileNotFoundError as e:
        print(f"❌ {e}")
        print(f"\n请复制 config_example.json 为 config.json 并配置您的飞书应用凭证")
    except Exception as e:
        print(f"❌ 错误: {e}")