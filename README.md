# Skills Demo：项目周报生成器（本地模型 + Ollama）

- 同一个模型（Ollama 本地模型）
- 通过开关加载不同 Skill（SOP + 模板 + 质检规则）
- 输出立刻从“聊天风”变成“公司可交付物风格”

## 1) 前置条件
1. 安装并启动 Ollama（本地）：https://ollama.com
2. 拉一个模型（示例）：
   - ollama pull qwen2.5:7b-instruct
   - 或者你本地已有的任意模型

## 2) 安装依赖
```bash
pip install -r requirements.txt
```

## 3) 运行
```bash
streamlit run app.py
```

## 4) 使用方式
1. 上传 data/weekly_tasks.xlsx（或你自己的 Excel/CSV）
2. 左侧选择 Skill：None / weekly_report
3. 点击“生成周报”对比效果

## 5) 数据列建议
project / module / task / owner / status / priority / due_date / progress / blocker / risk

如果你列名不同，tools.normalize_columns() 会做一些常见映射。
