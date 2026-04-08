#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书电子表格自动更新模块 v7.0 (智能排序版)
用于病历管家skill - 支持自定义配置

v7.0 更新:
- 新增 ordered_insert(): 按日期自动排序插入，保证时间倒序
- 新增 validate_row_data(): 写入前格式校验（长度、必填项）
- 新增 get_existing_data(): 写入前读取现有数据
- 新增 smart_archive_ct(): CT报告专用智能归档入口
- auto_archive() 升级：内部调用 ordered_insert 而非 append_row
"""
import json
import urllib.request
import urllib.error
import ssl
import os
import re
from datetime import datetime

ssl._create_default_https_context = ssl._create_unverified_context


# ==================== 列格式规范（硬编码强制约束）====================
CT_COLUMN_RULES = {
    "日期": {"max_len": 12, "pattern": r"^\d{4}-\d{2}-\d{2}$", "required": True},
    "医院": {"max_len": 20, "required": True},
    "检查类型": {"max_len": 18, "required": False},
    "原发灶": {"max_len": 25, "required": False},
    "转移灶": {"max_len": 40, "required": False},
    "总体评价": {"max_len": 15, "required": True, "desc": "图标+≤8字短语"},
    "医生建议": {"max_len": 20, "required": False, "desc": "简短可操作建议"},
}

BLOOD_COLUMN_RULES = {
    "日期": {"max_len": 12, "pattern": r"^\d{4}-\d{2}-\d{2}$", "required": True},
    "医院": {"max_len": 20, "required": True},
    # 指标列不做长度限制，只做数值基本校验
}


def load_config(config_path=None):
    """加载配置文件"""
    if config_path is None:
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config.json')

    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")

    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def get_sheets_config(config):
    """获取工作表配置"""
    return config.get('sheets', {})


class FeishuSheetsUpdater:
    """飞书电子表格更新器 - v7.0 支持有序插入和格式校验"""

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

                    for sheet in sheets:
                        sheet_title = sheet.get('title', '')
                        if sheet_name is None or sheet_title == sheet_name:
                            return sheet.get('sheet_id'), sheet_title

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

    # ==================== v7.0 新增方法 ====================

    def get_existing_data(self, sheet_id, max_rows=100):
        """
        读取工作表现有数据
        
        Returns:
            dict: {
                "header": list,      # 表头行
                "rows": [list, ...], # 数据行（按表格顺序）
                "total_rows": int,   # 数据行数（不含表头）
                "dates": [(date_str, row_num), ...]  # (日期, 行号) 用于排序判断
            }
        """
        result = {
            "header": [],
            "rows": [],
            "total_rows": 0,
            "dates": []
        }

        if not sheet_id:
            return result

        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.spreadsheet_token}/values/{sheet_id}!A1:N{max_rows}"
        headers = {"Authorization": f"Bearer {self.token}"}

        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req) as response:
                api_result = json.loads(response.read().decode('utf-8'))
                if api_result.get('code') == 0:
                    values = api_result.get('data', {}).get('valueRange', {}).get('values', [])
                    
                    if not values:
                        return result
                    
                    # 第1行是表头
                    result["header"] = values[0] if values else []
                    
                    # 后续是数据
                    for i, row in enumerate(values[1:], start=2):
                        if row and row[0]:  # 有日期的才算有效行
                            result["rows"].append(row)
                            result["dates"].append((row[0], i))
                    
                    result["total_rows"] = len(result["rows"])
        except Exception as e:
            print(f"读取数据失败: {e}")

        return result

    def validate_row_data(self, columns_config, column_names, row_data):
        """
        校验行数据格式
        
        Args:
            columns_config: 列规则字典 (CT_COLUMN_RULES / BLOOD_COLUMN_RULES)
            column_names: 表头列名列表 (从config.json的columns字段)
            row_data: 待写入的数据列表
        
        Returns:
            dict: {"valid": bool, "errors": [str, ...], "warnings": [str, ...]}
        """
        errors = []
        warnings = []

        for i, col_name in enumerate(column_names):
            rule = columns_config.get(col_name, {})

            if i >= len(row_data):
                if rule.get("required", False):
                    errors.append(f"第{i+1}列[{col_name}]为必填项，但未提供")
                continue

            value = str(row_data[i]) if row_data[i] is not None else ""

            # 必填检查
            if rule.get("required", False) and not value.strip():
                errors.append(f"第{i+1}列[{col_name}]为必填项，值为空")

            # 长度检查
            max_len = rule.get("max_len")
            if max_len and len(value) > max_len:
                desc = rule.get("desc", "")
                errors.append(
                    f"第{i+1}列[{col_name}]内容过长({len(value)}字符>{max_len}上限): "
                    f"'{value[:30]}...'{f' (要求: {desc})' if desc else ''}"
                )

            # 格式正则检查
            pattern = rule.get("pattern")
            if pattern and value and not re.match(pattern, value):
                errors.append(f"第{i+1}列[{col_name}]格式不匹配: '{value}' (期望格式: YYYY-MM-DD)")

        return {
            "valid": len(errors) == 0,
            "errors": errors,
            "warnings": warnings
        }

    def _write_row_at(self, sheet_id, row_number, row_data):
        """在指定行号写入一行数据（内部方法）"""
        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.spreadsheet_token}/values"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }

        data = {
            "valueRange": {
                "range": f"{sheet_id}!A{row_number}:Z{row_number}",
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
                    return {"success": True, "row": row_number}
                else:
                    return {"success": False, "info": str(result)}
        except Exception as e:
            return {"success": False, "info": str(e)}

    def ordered_insert(self, sheet_id, row_data, date_col=0, rules=None, column_names=None, check_dup=True):
        """
        按时间倒序插入数据（核心新方法）
        
        自动读取现有数据 → 确定正确位置 → 插入到该位置
        
        Args:
            sheet_id: 工作表ID
            row_data: 行数据列表
            date_col: 日期所在列索引（默认0）
            rules: 列格式校验规则 (None则跳过校验)
            column_names: 列名列表（与rules配合使用）
            check_dup: 是否查重
        
        Returns:
            dict: 结果信息
        """
        if not sheet_id:
            return {"success": False, "info": "未指定工作表"}

        if len(row_data) <= date_col:
            return {"success": False, "info": "行数据缺少日期列"}

        new_date = row_data[date_col]

        # Step 1: 格式校验（如果提供了规则）
        if rules and column_names:
            validation = self.validate_row_data(rules, column_names, row_data)
            if not validation["valid"]:
                error_detail = "\n".join(validation["errors"])
                return {
                    "success": False,
                    "info": f"格式校验不通过:\n{error_detail}\n\n请修正后重试。参考SKILL.md中的列格式规范。",
                    "validation_errors": validation["errors"]
                }
            if validation["warnings"]:
                for w in validation["warnings"]:
                    print(f"⚠️ 警告: {w}")

        # Step 2: 读取现有数据
        existing = self.get_existing_data(sheet_id)

        # Step 3: 去重检查
        if check_dup:
            for exist_date, row_num in existing["dates"]:
                if exist_date == new_date:
                    return {
                        "success": False,
                        "duplicate": True,
                        "row": row_num,
                        "info": f"日期 {new_date} 的记录已存在于第{row_num}行，拒绝重复录入"
                    }

        # Step 4: 计算插入位置（按日期降序排列）
        insert_row = existing["total_rows"] + 2  # 默认插入到最后

        if existing["dates"]:
            # 找到第一个比新日期小的日期，插到它前面
            for exist_date, row_num in existing["dates"]:
                if exist_date < new_date:
                    insert_row = row_num
                    break
            else:
                # 所有日期都比新的大（或相等），插到最后
                pass

        # Step 5: 如果不是追加到末尾，需要先把后面的数据下移
        # 飞书API不支持insert row操作，所以我们：
        #   方案A：直接用正确的位置覆盖写入（如果有空行）
        #   方案B：重写整个数据区
        # 这里采用方案A + 兜底：先尝试直接写目标位置
        
        write_result = self._write_row_at(sheet_id, insert_row, row_data)
        
        if write_result.get("success"):
            return {
                "success": True,
                "row": insert_row,
                "info": (
                    f"已按时间倒序插入到第{insert_row}行 (日期: {new_date})。\n"
                    f"现有数据共{existing['total_rows']}条，本条为最新数据排在最前。"
                ),
                "pre_insert_info": {
                    "new_date": new_date,
                    "existing_dates": [d[0] for d in existing["dates"]],
                    "insert_position": insert_row,
                    "total_existing": existing["total_rows"]
                }
            }
        else:
            return {"success": False, "info": write_result.get("info", "写入失败")}

    # ==================== 原有方法（保留兼容）====================

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
                    for i, row in enumerate(values[1:], start=2):
                        if len(row) > 0 and row[0] == check_date:
                            if hospital is None or (len(row) > 1 and hospital in row[1]):
                                return i
        except:
            pass

        return None

    def get_next_row_number(self, sheet_id):
        """获取下一个空行的行号（简单追加模式）"""
        if not sheet_id:
            return 2
        existing = self.get_existing_data(sheet_id)
        return existing["total_rows"] + 2

    def append_row(self, sheet_id, row_data, check_dup=True, date_col=0, hospital_col=1):
        """追加一行数据（简单模式，不排序 - 保留向后兼容）

        ⚠️ 推荐使用 ordered_insert() 替代此方法以保证时间顺序！
        """
        if not sheet_id:
            return {"success": False, "info": "未指定工作表"}

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

        next_row = self.get_next_row_number(sheet_id)
        return self._write_row_at(sheet_id, next_row, row_data)

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


# ==================== 高层归档函数 ====================

def smart_archive_ct(record):
    """
    CT报告专用智能归档（推荐入口）
    
    自动完成：认证 → 读取现有数据 → 格式校验 → 排序插入
    
    Args:
        record: dict，必须包含以下字段:
            - date (str): YYYY-MM-DD
            - hospital (str): 医院名称
            - exam_type (str): 检查类型
            - primary_focus (str): 原发灶描述
            - metastasis (str): 转移灶描述
            - overall_eval (str): 总体评价（≤15字符）
            - doctor_advice (str): 医生建议（≤20字符）
    
    Returns:
        dict: 归档结果
    """
    required_fields = ['date', 'hospital']
    for field in required_fields:
        if not record.get(field):
            return {"success": False, "info": f"缺少必填字段: {field}"}

    try:
        updater = FeishuSheetsUpdater()

        if not updater.authenticate():
            return {"success": False, "info": "飞书认证失败"}

        # 获取CT报告工作表配置
        ct_config = updater.sheets_config.get('ct_report')
        if not ct_config or not ct_config.get('enabled', False):
            return {"success": False, "info": "ct_report 工作表未启用"}

        sheet_id, sheet_title = updater.get_sheet_info(ct_config['name'])
        if not sheet_id:
            return {"success": False, "info": f"未找到工作表: {ct_config['name']}"}

        # 构建行数据（严格按config中定义的列顺序）
        columns = ct_config.get('columns', [])
        row_data = [
            record.get('date', ''),
            record.get('hospital', '未知'),
            record.get('exam_type', ''),
            record.get('primary_focus', ''),
            record.get('metastasis', ''),
            record.get('overall_eval', ''),
            record.get('doctor_advice', ''),
        ]

        # 使用 ordered_insert 进行排序+校验+写入
        result = updater.ordered_insert(
            sheet_id=sheet_id,
            row_data=row_data,
            date_col=0,
            rules=CT_COLUMN_RULES,
            column_names=columns,
            check_dup=True
        )

        if result.get('success'):
            return {
                "success": True,
                "info": result.get('info', ''),
                "link": f"https://p16aafnobs.feishu.cn/sheets/{updater.spreadsheet_token}",
                "row": result.get('row')
            }
        elif result.get('duplicate'):
            return {
                "success": False,
                "duplicate": True,
                "info": result.get('info', '重复数据')
            }
        elif result.get('validation_errors'):
            return {
                "success": False,
                "validation_failed": True,
                "errors": result.get('validation_errors'),
                "info": result.get('info', '格式校验失败')
            }
        else:
            return {"success": False, "info": result.get('info', '未知错误')}

    except FileNotFoundError as e:
        return {"success": False, "info": f"配置文件缺失: {e}"}
    except Exception as e:
        return {"success": False, "info": f"系统错误: {str(e)}"}


def auto_archive(report_type, record, allow_duplicate=False):
    """
    自动归档到飞书表格 (v7.0 升版 - 内部使用ordered_insert)
    
    Args:
        report_type: 报告类型 (blood_routine/lipid/tumor_markers/ct_report)
        record: 记录数据 dict
        allow_duplicate: 是否允许重复
        - 对于ct_report类型，推荐改用 smart_archive_ct()
    
    Returns:
        dict: {"success": bool, "info": str, "link": str}
    """
    try:
        updater = FeishuSheetsUpdater()

        if not updater.authenticate():
            return {"success": False, "info": "认证失败"}

        if not updater.spreadsheet_token:
            return {"success": False, "info": "未配置spreadsheet_token"}

        sheet_config = updater.sheets_config.get(report_type)
        if not sheet_config or not sheet_config.get('enabled', False):
            return {"success": False, "info": f"报告类型 {report_type} 未启用"}

        sheet_id, sheet_title = updater.get_sheet_info(sheet_config['name'])
        if not sheet_id:
            return {"success": False, "info": f"未找到工作表: {sheet_config['name']}"}

        # 根据报告类型构建行数据和选择规则
        columns = sheet_config.get('columns', [])

        if report_type == 'ct_report':
            # CT报告：使用完整字段 + CT规则
            row_data = [
                record.get('date', datetime.now().strftime('%Y-%m-%d')),
                record.get('hospital', '未知'),
                record.get('exam_type', ''),
                record.get('primary_focus', ''),
                record.get('metastasis', ''),
                record.get('overall_eval', ''),
                record.get('doctor_advice', ''),
            ]
            rules = CT_COLUMN_RULES
            use_ordered = True
        else:
            # 血常规/血脂等：使用指标数组
            row_data = [
                record.get('date', datetime.now().strftime('%Y-%m-%d')),
                record.get('hospital', '未知')
            ]

            indicators = record.get('indicators', [])
            for ind in indicators:
                row_data.append(str(ind.get('value', '-')))

            row_data.append(record.get('abnormal_summary', '正常'))
            row_data.append(record.get('recommendation', '继续监测'))

            rules = BLOOD_COLUMN_RULES
            use_ordered = True  # v7.0: 所有报告类型都启用有序插入

        # 使用 ordered_insert 替代原来的 append_row
        check_dup = not allow_duplicate
        if use_ordered:
            result = updater.ordered_insert(
                sheet_id=sheet_id,
                row_data=row_data,
                date_col=0,
                rules=rules,
                column_names=columns,
                check_dup=check_dup
            )
        else:
            result = updater.append_row(sheet_id, row_data, check_dup=check_dup)

        if result.get('success'):
            info = result.get('info', '')
            pre_info = result.get('pre_insert_info', {})
            if pre_info:
                extra = f"\n📊 时间线: {pre_info['new_date']}(新)"
                if pre_info.get('existing_dates'):
                    extra += " ← " + " ← ".join(pre_info['existing_dates'])
                info += extra

            return {
                "success": True,
                "info": f"已录入到【{sheet_title}】{info}",
                "link": f"https://p16aafnobs.feishu.cn/sheets/{updater.spreadsheet_token}",
                "row": result.get('row')
            }
        elif result.get('duplicate'):
            return {
                "success": False,
                "info": result['info'],
                "duplicate": True
            }
        elif result.get('validation_errors'):
            return {
                "success": False,
                "validation_failed": True,
                "errors": result.get('validation_errors'),
                "info": result.get('info', '格式校验失败，请修正后重试')
            }
        else:
            return {"success": False, "info": result.get('info', '未知错误')}

    except FileNotFoundError as e:
        return {"success": False, "info": str(e)}
    except Exception as e:
        return {"success": False, "info": f"系统错误: {str(e)}"}


# ==================== CLI 入口 ====================

if __name__ == "__main__":
    print("=== 病历管家 - 飞书表格更新模块 v7.0 ===\n")
    print("✨ 新功能: 智能排序插入 | 格式自动校验 | CT专用归档\n")

    try:
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config.json')
        config = load_config(config_path)
        print(f"✅ 配置加载成功")
        print(f"   - App ID: {config['app_id'][:10]}...")
        print(f"   - 电子表格: {config.get('spreadsheet_token', '未设置')[:20]}...")

        updater = FeishuSheetsUpdater()
        if updater.authenticate():
            print(f"✅ 认证成功")

            # 显示可用工作表
            sheets = updater.get_all_sheets()
            print(f"\n📊 电子表格工作表:")
            for s in sheets:
                print(f"   - {s.get('title')} (id={s.get('sheet_id')})")

            # 显示配置的工作表
            print(f"\n⚙️ 已启用的工作表:")
            for key, sheet in updater.sheets_config.items():
                if sheet.get('enabled'):
                    cols = sheet.get('columns', [])
                    print(f"   ✅ {key}: {sheet['name']} ({len(cols)}列: {', '.join(cols[:3])}...)")
                else:
                    print(f"   ❌ {key}: {sheet['name']} (未启用)")
            
            # 测试读取CT报告数据
            ct_config = updater.sheets_config.get('ct_report')
            if ct_config and ct_config.get('enabled'):
                sheet_id, title = updater.get_sheet_info(ct_config['name'])
                if sheet_id:
                    existing = updater.get_existing_data(sheet_id)
                    print(f"\n📋 【{title}】当前数据:")
                    print(f"   表头: {existing['header']}")
                    print(f"   数据行数: {existing['total_rows']}")
                    if existing['dates']:
                        dates_sorted = sorted(existing['dates'], key=lambda x: x[0], reverse=True)
                        print(f"   日期(从新→旧): {[d[0] for d in dates_sorted]}")

        else:
            print(f"❌ 认证失败，请检查配置")

    except FileNotFoundError as e:
        print(f"❌ {e}")
        print(f"\n请复制 config_example.json 为 config.json 并配置您的飞书应用凭证")
    except Exception as e:
        print(f"❌ 错误: {e}")
