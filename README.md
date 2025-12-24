# 简历解析与PPT生成工具

通过 Coze AI 工作流自动解析简历文件，提取结构化信息并生成专业的 PPT 演示文稿。

## 功能

- 支持 PDF、DOCX 格式简历
- AI 自动提取个人信息、论文、项目、获奖等内容
- 论文以表格形式分页展示
- 可同时处理多份简历

## 快速开始

### 1. 安装依赖

```bash
pip install requests python-pptx
```

### 2. 配置

复制配置模板并填写：

```bash
cp config.example.json config.json
```

编辑 `config.json`：

```json
{
    "token": "your_coze_api_token",
    "workflow_id": "your_workflow_id",
    "local_files": [
        "path/to/resume1.pdf",
        "path/to/resume2.docx"
    ]
}
```

> 注意：`config.json` 已加入 `.gitignore`，不会被提交到版本库

### 3. 运行

```bash
# 完整流程: 上传文件 → 调用API → 生成PPT
python run.py

# 测试模式: 从本地数据生成PPT（无需API）
python run.py --from-file response.txt
```

## 文件说明

| 文件 | 说明 |
|------|------|
| `run.py` | 主程序，处理文件上传和API调用 |
| `ppt_renderer.py` | PPT渲染模块，生成演示文稿 |
| `config.json` | 配置文件（敏感信息，不提交） |
| `config.example.json` | 配置模板 |
| `response.txt` | 测试用API响应数据 |

## PPT 内容

生成的 PPT 包含以下页面：

1. **封面** - 姓名、学历、专业
2. **个人信息** - 基本信息 + 教育背景
3. **论文成果** - 表格展示，自动分页
4. **项目情况** - 获批项目列表
5. **获奖成果** - 获奖、专利、著作

## 自定义

### 调整每页论文数量

```python
from ppt_renderer import PPTRenderer

renderer = PPTRenderer(papers_per_page=10)
renderer.render_all(resumes)
```

### 修改配色

编辑 `ppt_renderer.py` 中的 `ColorScheme` 类：

```python
class ColorScheme:
    PRIMARY = RgbColor(0x0F, 0x4C, 0x81)    # 主色调
    ACCENT_GOLD = RgbColor(0xD4, 0xAF, 0x37) # 强调色
    # ...
```

## 注意事项

- Coze Token 有效期有限，过期需重新获取
- 论文"一作/通讯"字段需手动补充
- 建议检查生成的 PPT 内容是否完整

## License

MIT
