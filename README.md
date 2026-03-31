# 月报竞品新闻渠道统计分析工具

## 功能说明

- ✅ 支持上传 Excel 文件
- ✅ 实时显示处理进度条
- ✅ 自动统计品牌声量 TOP10 媒体

统计月报重点竞品新闻渠道数据，包括：
- 按品牌统计声量（篇）TOP10 媒体
- 触达（人次）以最新日期数据为准
- 自动识别媒体类别（汽车行业垂类媒体、商业类媒体、综合类媒体、生活类媒体）

## 本地使用

### 命令行版本

```bash
cd d:/trae/YBZDJPQD
./venv/Scripts/python competitor_news_analyzer.py <输入文件.xlsx>
```

### Web 版本

```bash
cd d:/trae/YBZDJPQD
./venv/Scripts/python app.py
# 访问 http://localhost:5000
```

## 部署到 Render

### 1. 准备 GitHub 仓库

```bash
cd d:/trae/YBZDJPQD
git init
git add .
git commit -m "feat: 月报竞品新闻渠道统计分析工具"
git branch -M main
git remote add origin https://github.com/你的用户名/competitor-news-analyzer.git
git push -u origin main
```

### 2. 在 Render 上部署

1. 登录 [Render](https://render.com/)
2. 点击 **New +** → **Web Service**
3. 关联你的 GitHub 仓库
4. 配置以下内容：

| 设置项 | 值 |
|--------|-----|
| Name | `competitor-news-analyzer`（或其他名称） |
| Region | Singapore（就近选择） |
| Branch | `main` |
| Runtime | `Python 3` |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `gunicorn app:app --bind 0.0.0.0:$PORT` |
| Plan | Free |

5. 点击 **Create Web Service**

### 3. 等待部署完成

部署成功后，Render 会提供一个 URL，例如：
`https://competitor-news-analyzer.onrender.com`

任何人访问此链接都可以使用工具上传 Excel 文件并获取分析结果。

## 输入表格格式

| Media ID | Date | Source | Potential Audience |
|----------|------|--------|-------------------|
| 品牌名称 | 日期 | 媒体名称 | 触达人数 |

## 输出表格格式

| 序号 | 品牌名称 | 媒体名称 | 声量（篇） | 触达（人次） | 媒体类别 |
|------|----------|----------|------------|--------------|----------|
| 1 | 比亚迪 | 汽车之家 | 25 | 1500.5K | 汽车行业垂类媒体 |

## 媒体类别规则

工具自动识别以下4类媒体：

1. **汽车行业垂类媒体**：汽车之家、懂车帝、易车、Bil24、Elbil24.no 等
2. **商业类媒体**：36氪、虎嗅、钛媒体、E24.no、Finansavisen 等
3. **综合类媒体**：人民日报、新华社、TV2 Nyheter、VG Online 等
4. **生活类媒体**：微信、微博、小红书、抖音 等

如需添加新媒体的类别映射，请编辑 `app.py` 或 `competitor_news_analyzer.py` 中的 `MEDIA_CATEGORIES` 字典。

## 项目文件结构

```
YBZDJPQD/
├── venv/                          # Python 虚拟环境
├── app.py                         # Flask Web 应用
├── competitor_news_analyzer.py    # 命令行版本
├── requirements.txt               # Python 依赖
├── Procfile                       # Render 部署配置
├── README.md                      # 使用说明
└── .gitignore                     # Git忽略文件
```
