# 病历管家 Skill 分享指南

## 目录

1. [分享架构设计](#1-分享架构设计)
2. [GitHub 公开仓库部署](#2-github-公开仓库部署)
3. [配置说明](#3-配置说明)
4. [飞书集成配置](#4-飞书集成配置)
5. [自定义报告类型](#5-自定义报告类型)
6. [常见问题](#6-常见问题)

---

## 1. 分享架构设计

### 架构方案：中心化云同步

```
管理员
    ↓ 迭代优化Skill
    ↓ 上传到云端
云端Skill仓库
    ↓ 自动同步
用户
    ↓ 自动获取更新
    ↓ 使用最新版本
```

---

## 2. GitHub 公开仓库部署

### 方案1: GitHub公开仓库(推荐) ⭐

**优点**:
- ✅ 完全免费
- ✅ 版本控制清晰
- ✅ 自动更新机制
- ✅ 社区协作方便
- ✅ 文档托管完善

**实施步骤**:

1. **创建GitHub仓库**
```bash
仓库名: medical-record-butler-skill
可见性: Public (公开)
```

2. **目录结构**
```
medical-record-butler-skill/
├── README.md                    # 使用说明
├── config_example.json           # 配置示例（公开，无敏感信息）
├── LICENSE                       # MIT许可证
├── .gitignore                    # 忽略敏感文件
├── skill/                       # Skill核心文件
│   ├── SKILL.md
│   ├── sharing_guide.md
│   ├── scripts/
│   │   └── feishu_api.py        # 飞书API脚本
│   └── references/
│       ├── blood_test_reference.json
│       ├── status_standards.md
│       └── FORMAT_VALIDATION_RULES.md
└── docs/                        # 文档
    ├── 用户指南.md
    └── 配置说明.md
```

3. **自动更新机制**
- 每次更新后发布新版本
- 用户CodeBuddy自动检测更新
- 提示用户安装新版本

---

## 3. 配置说明

### 配置文件结构

复制 `config_example.json` 为 `config.json`，填入您的飞书应用凭证：

```json
{
  "app_id": "您的飞书应用ID",
  "app_secret": "您的飞书应用密钥",
  "spreadsheet_token": "您的电子表格Token",
  "sheets": {
    "blood_routine": {
      "name": "血常规",
      "enabled": true,
      "columns": ["日期", "医院", "WBC", "ANC", "PLT", "Hb", "异常", "建议"]
    },
    "lipid": {
      "name": "血脂",
      "enabled": false,
      "columns": ["日期", "医院", "TC", "TG", "HDL-C", "LDL-C", "异常", "建议"]
    },
    "tumor_markers": {
      "name": "肿瘤标志物",
      "enabled": false,
      "columns": ["日期", "医院", "CA-125", "CYFRA21-1", "VEGF", "SCC", "异常", "建议"]
    },
    "ct_report": {
      "name": "CT报告",
      "enabled": false,
      "columns": ["日期", "医院", "检查类型", "原发灶", "转移灶", "总体评价", "医生建议"]
    }
  },
  "report_types": {
    "blood_routine": {
      "name": "血常规",
      "keywords": ["血常规", "血分析", "Blood", "WBC", "白细胞"],
      "indicators": ["WBC", "ANC", "PLT", "Hb", "CRP", "LYM"]
    },
    "lipid": {
      "name": "血脂",
      "keywords": ["血脂", "胆固醇", "甘油三酯", "Lipid"],
      "indicators": ["TC", "TG", "HDL-C", "LDL-C"]
    },
    "tumor_markers": {
      "name": "肿瘤标志物",
      "keywords": ["肿瘤", "CEA", "CA-125", "CA19-9", "Tumor"],
      "indicators": ["CA-125", "CYFRA21-1", "VEGF", "SCC", "CEA", "CA19-9"]
    },
    "ct_report": {
      "name": "CT/MRI",
      "keywords": ["CT", "MRI", "影像", "CT报告", "核磁"],
      "indicators": []
    }
  },
  "settings": {
    "deduplication": {
      "enabled": true,
      "key_fields": ["date", "hospital", "report_type"]
    },
    "auto_trend": true,
    "emergency_warning": true
  }
}
```

### 获取飞书配置

1. **创建飞书应用**: 前往 [飞书开放平台](https://open.feishu.cn/) 创建应用
2. **获取App ID和App Secret**: 在应用详情页获取
3. **获取Spreadsheet Token**: 
   - 创建或打开一个飞书电子表格
   - 从URL中提取Token: `https://p16aafnobs.feishu.cn/sheets/{TOKEN}`
   - Token格式: `xxxxxxxxxxxxxxxx`

---

## 4. 飞书集成配置

### 权限要求

飞书应用需要以下权限：
- `sheets:spreadsheet:readonly` - 读取电子表格
- `sheets:spreadsheet:write` - 写入电子表格
- `sheets:sheet:readonly` - 读取工作表
- `sheets:value:write` - 写入单元格

### 配置步骤

1. 在飞书开放平台创建应用
2. 添加所需权限
3. 发布版本并安装到租户
4. 获取 App ID 和 App Secret
5. 创建或使用现有电子表格
6. 复制 Token 到配置文件

---

## 5. 自定义报告类型

### 添加新的报告类型

用户可以在 `config.json` 中添加自定义报告类型：

```json
{
  "sheets": {
    "new_report_type": {
      "name": "新报告类型",
      "enabled": true,
      "columns": ["日期", "医院", "指标1", "指标2", "异常", "建议"]
    }
  },
  "report_types": {
    "new_report_type": {
      "name": "新报告",
      "keywords": ["关键词1", "关键词2"],
      "indicators": ["指标1", "指标2"]
    }
  }
}
```

### 配置说明

| 配置项 | 说明 |
|--------|------|
| `name` | 工作表显示名称 |
| `enabled` | 是否启用此报告类型 |
| `columns` | 列标题（按顺序） |
| `keywords` | 识别报告类型的关键词 |
| `indicators` | 需要提取的指标列表 |

---

## 6. 常见问题

### Q1: 如何只使用部分报告类型？

在 `config.json` 中设置 `enabled: false` 即可禁用不需要的报告类型。

### Q2: 可以使用自己的电子表格吗？

可以。只需将您的 `spreadsheet_token` 填入配置文件即可。

### Q3: 如何添加更多工作表？

在 `sheets` 配置中添加新的工作表配置，并确保电子表格中存在对应的工作表。

### Q4: 配置文件中可以包含中文吗？

可以。JSON 文件支持 UTF-8 编码，中文内容会正常显示。

### Q5: 如何启用数据去重？

默认已启用。在 `settings.deduplication.enabled` 中可以控制开关。

---

## 安全注意事项

1. **不要提交敏感信息**: `config.json` 包含凭据，应加入 `.gitignore`
2. **使用环境变量**: 生产环境建议使用环境变量存储敏感信息
3. **定期更新密钥**: 定期更换飞书应用密钥以确保安全

---

## 技术支持

如有问题，请提交 Issue 或查看文档。