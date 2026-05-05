# 云端自动化测试

基于 Selenium 和 Flask 的云端自动化测试平台。

## 功能特性

- 自动化浏览器操作
- Web 应用测试
- 截图功能
- 日志记录
- Flask Web 界面

## 技术栈

- Python
- Selenium
- Flask
- OpenAI Agents SDK

## 安装

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 配置环境变量：
创建 `.env` 文件并设置必要的环境变量（如 OPENAI_API_KEY）

## 运行

```bash
python app.py
```

访问 http://localhost:8080 使用 Web 界面

## 项目结构

- `app.py` - Flask 主应用
- `selenium_tools.py` - Selenium 自动化工具
- `agent_config.py` - Agent 配置
- `templates/` - HTML 模板
- `log/` - 日志文件（不被 Git 跟踪）
- `screenshots/` - 截图文件（不被 Git 跟踪）

## 注意事项

- `log/` 和 `screenshots/` 目录已被添加到 `.gitignore`，不会被提交到 Git 仓库
- 敏感信息应存储在 `.env` 文件中，该文件也不会被提交
