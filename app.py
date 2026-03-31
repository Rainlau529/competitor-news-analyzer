"""
月报竞品新闻渠道统计分析工具 - Web服务
支持文件上传，自动处理并返回结果，带进度条显示
"""

import os
import pandas as pd
from datetime import datetime
from flask import Flask, request, send_file, render_template_string, Response
from io import BytesIO

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'competitor-news-analyzer-secret-key')

# 媒体类别映射表
MEDIA_CATEGORIES = {
    # ========== 汽车行业垂类媒体 ==========
    "汽车之家": "汽车行业垂类媒体",
    "懂车帝": "汽车行业垂类媒体",
    "易车": "汽车行业垂类媒体",
    "太平洋汽车": "汽车行业垂类媒体",
    "爱卡汽车": "汽车行业垂类媒体",
    "网易汽车": "汽车行业垂类媒体",
    "新浪汽车": "汽车行业垂类媒体",
    "腾讯汽车": "汽车行业垂类媒体",
    "凤凰汽车": "汽车行业垂类媒体",
    "搜狐汽车": "汽车行业垂类媒体",
    "车质网": "汽车行业垂类媒体",
    "中国汽车报": "汽车行业垂类媒体",
    "autohome": "汽车行业垂类媒体",
    "dongchedi": "汽车行业垂类媒体",
    "yiche": "汽车行业垂类媒体",
    "Bil24": "汽车行业垂类媒体",
    "BilNorge.no": "汽车行业垂类媒体",
    "Bilnytt.no": "汽车行业垂类媒体",
    "Nybiltester": "汽车行业垂类媒体",
    "Nyheter fra Bilglede": "汽车行业垂类媒体",
    "Magasinet Motor": "汽车行业垂类媒体",
    "Elbil24.no": "汽车行业垂类媒体",
    "OtonomHaber (NO)": "汽车行业垂类媒体",

    # ========== 商业类媒体 ==========
    "第一财经": "商业类媒体",
    "每日经济新闻": "商业类媒体",
    "经济观察报": "商业类媒体",
    "36氪": "商业类媒体",
    "虎嗅": "商业类媒体",
    "钛媒体": "商业类媒体",
    "财经网": "商业类媒体",
    "华尔街见闻": "商业类媒体",
    "界面新闻": "商业类媒体",
    "雪球": "商业类媒体",
    "东方财富": "商业类媒体",
    "E24.no": "商业类媒体",
    "Finansavisen": "商业类媒体",
    "Offshore.no": "商业类媒体",
    "Teknisk Ukeblad Online": "商业类媒体",
    "ITavisen.no": "商业类媒体",
    "Lyd & Bilde Online": "商业类媒体",
    "Tek.no": "商业类媒体",
    "Gamereactor (NO)": "商业类媒体",
    "Teksiden": "商业类媒体",

    # ========== 综合类媒体 ==========
    "人民日报": "综合类媒体",
    "新华社": "综合类媒体",
    "央视": "综合类媒体",
    "人民网": "综合类媒体",
    "新华网": "综合类媒体",
    "澎湃新闻": "综合类媒体",
    "今日头条": "综合类媒体",
    "腾讯新闻": "综合类媒体",
    "网易新闻": "综合类媒体",
    "新浪新闻": "综合类媒体",
    "搜狐新闻": "综合类媒体",
    "凤凰资讯": "综合类媒体",
    "TV2 Nyheter": "综合类媒体",
    "Dagbladet Online": "综合类媒体",
    "VG Online": "综合类媒体",
    "Nordlys.no": "综合类媒体",
    "The News (NO)": "综合类媒体",
    "Norway News (NO)": "综合类媒体",
    "Vestlandsnytt": "综合类媒体",
    "Dagens (NO)": "综合类媒体",
    "MSN": "综合类媒体",

    # ========== 生活类媒体 ==========
    "微信": "生活类媒体",
    "微博": "生活类媒体",
    "小红书": "生活类媒体",
    "抖音": "生活类媒体",
    "快手": "生活类媒体",
    "Bilibili": "生活类媒体",
    "知乎": "生活类媒体",
    "百度": "生活类媒体",
    "UC": "生活类媒体",
    "Dinside.no": "生活类媒体",
    "Bodonu": "生活类媒体",
}


def format_audience(value):
    """格式化触达人数，添加单位，不四舍五入"""
    if pd.isna(value) or value == 0:
        return "-"
    if value >= 1_000_000:
        result = value / 1_000_000
        integer_part = int(result)
        decimal_part = int((result - integer_part) * 10)
        return f"{integer_part}.{decimal_part}M"
    elif value >= 1_000:
        result = value / 1_000
        integer_part = int(result)
        decimal_part = int((result - integer_part) * 10)
        return f"{integer_part}.{decimal_part}K"
    else:
        return str(int(value))


def get_media_category(media_name):
    """获取媒体类别"""
    if pd.isna(media_name):
        return "综合类媒体"
    media_name_str = str(media_name).strip()
    if media_name_str in MEDIA_CATEGORIES:
        return MEDIA_CATEGORIES[media_name_str]
    for key, category in MEDIA_CATEGORIES.items():
        if key in media_name_str or media_name_str in key:
            return category
    return "综合类媒体"


def analyze_competitor_news_stream(df, progress_callback=None):
    """分析竞品新闻渠道数据，支持进度回调"""

    def send_progress(step, message, percent):
        if progress_callback:
            progress_callback(step, message, percent)

    send_progress(1, "正在检查数据格式...", 10)

    # 检查必要的列
    required_columns = ["Media ID", "Date", "Source", "Potential Audience"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        return None, f"缺少必要的列: {', '.join(missing_columns)}"

    send_progress(2, "正在转换日期格式...", 20)

    # 转换日期列
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # 获取唯一品牌
    brands = df["Media ID"].dropna().unique().tolist()
    total_brands = len(brands)

    send_progress(3, f"发现 {total_brands} 个品牌，开始分析...", 30)

    # 预计算每个(品牌, 媒体)的最新触达数
    latest_audience_cache = {}
    brand_idx = 0
    for brand in brands:
        brand_idx += 1
        brand_df = df[df["Media ID"] == brand].copy()
        send_progress(3, f"正在分析品牌 {brand_idx}/{total_brands}：{brand}...", 30 + int(brand_idx / total_brands * 40))

        for source in brand_df["Source"].unique():
            media_data = brand_df[brand_df["Source"] == source].copy()
            media_data = media_data.dropna(subset=["Date"])
            if media_data.empty:
                latest_audience_cache[(brand, source)] = None
            else:
                latest_idx = media_data["Date"].idxmax()
                latest_row = media_data.loc[latest_idx]
                latest_audience_cache[(brand, source)] = latest_row["Potential Audience"]

    send_progress(4, "正在生成TOP10排名...", 75)

    # 存储所有结果
    all_results = []

    for i, brand in enumerate(brands):
        brand_df = df[df["Media ID"] == brand].copy()

        # 统计每个媒体的声量（篇数）
        volume_df = brand_df.groupby("Source").size().reset_index(name="声量（篇）")

        # 获取每个媒体最新日期的触达数
        def get_latest_audience(media_name):
            return latest_audience_cache.get((brand, media_name), None)

        volume_df["触达（人次）"] = volume_df["Source"].apply(get_latest_audience)

        # 按声量降序排列，取TOP10
        volume_df = volume_df.sort_values("声量（篇）", ascending=False).head(10)
        volume_df["品牌名称"] = brand
        volume_df["媒体类别"] = volume_df["Source"].apply(get_media_category)
        volume_df["触达（人次）"] = volume_df["触达（人次）"].apply(format_audience)
        volume_df.insert(0, "序号", range(1, len(volume_df) + 1))
        volume_df = volume_df.rename(columns={"Source": "媒体名称"})
        volume_df = volume_df[["序号", "品牌名称", "媒体名称", "声量（篇）", "触达（人次）", "媒体类别"]]

        all_results.append(volume_df)

    send_progress(5, "正在生成结果文件...", 90)

    return all_results, None


# HTML模板
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>月报竞品新闻渠道统计分析</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 40px 20px;
        }
        .container { max-width: 900px; margin: 0 auto; }
        .card {
            background: white;
            border-radius: 16px;
            padding: 40px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }
        h1 { color: #333; text-align: center; margin-bottom: 10px; font-size: 28px; }
        .subtitle { text-align: center; color: #666; margin-bottom: 30px; font-size: 14px; }

        .upload-area {
            border: 2px dashed #ddd;
            border-radius: 12px;
            padding: 40px;
            text-align: center;
            margin-bottom: 20px;
            transition: all 0.3s;
            cursor: pointer;
        }
        .upload-area:hover { border-color: #667eea; background: #f8f9ff; }
        .upload-area.dragover { border-color: #667eea; background: #f0f3ff; }
        .upload-icon { font-size: 48px; margin-bottom: 15px; }
        .upload-text { color: #666; font-size: 16px; }
        .upload-text span { color: #667eea; font-weight: 600; }
        input[type="file"] { display: none; }

        .btn {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 14px 32px;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        .btn:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(102, 126, 234, 0.4); }
        .btn:disabled { background: #ccc; cursor: not-allowed; transform: none; box-shadow: none; }
        .submit-area { text-align: center; margin-top: 20px; }

        /* 进度条样式 */
        .progress-container {
            display: none;
            margin: 20px 0;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 12px;
        }
        .progress-container.active { display: block; }
        .progress-status {
            font-size: 14px;
            color: #666;
            margin-bottom: 10px;
            text-align: center;
        }
        .progress-bar-bg {
            width: 100%;
            height: 12px;
            background: #e9ecef;
            border-radius: 6px;
            overflow: hidden;
        }
        .progress-bar {
            width: 0%;
            height: 100%;
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
            border-radius: 6px;
            transition: width 0.3s ease;
        }
        .progress-percent {
            font-size: 24px;
            font-weight: 600;
            color: #667eea;
            text-align: center;
            margin-top: 10px;
        }

        .note {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 16px;
            margin-top: 20px;
            font-size: 13px;
            color: #666;
        }
        .note h3 { color: #333; margin-bottom: 8px; font-size: 14px; }
        .note ul { margin-left: 20px; }

        .error {
            background: #f8d7da;
            color: #721c24;
            padding: 16px;
            border-radius: 8px;
            margin-bottom: 20px;
        }

        .result-info {
            background: #d4edda;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 20px;
            color: #155724;
        }

        .brand-section { margin-bottom: 30px; }
        .brand-title {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 10px 20px;
            border-radius: 8px 8px 0 0;
            font-weight: 600;
        }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th {
            background: #f8f9fa;
            padding: 12px 8px;
            text-align: center;
            font-weight: 600;
            color: #333;
            border-bottom: 2px solid #dee2e6;
        }
        td {
            padding: 10px 8px;
            text-align: center;
            border-bottom: 1px solid #dee2e6;
        }
        tr:hover { background: #f8f9fa; }
        .category-tag {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 500;
        }
        .category-汽车 { background: #e3f2fd; color: #1565c0; }
        .category-商业 { background: #fff3e0; color: #e65100; }
        .category-综合 { background: #f3e5f5; color: #7b1fa2; }
        .category-生活 { background: #e8f5e9; color: #2e7d32; }

        @media (max-width: 600px) {
            .card { padding: 20px; }
            table { font-size: 12px; }
            th, td { padding: 8px 4px; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="card">
            <h1>月报竞品新闻渠道统计分析</h1>
            <p class="subtitle">上传Excel文件，自动统计品牌声量TOP10媒体</p>

            <form id="uploadForm">
                <div class="upload-area" id="uploadArea">
                    <div class="upload-icon">📁</div>
                    <p class="upload-text">
                        点击或拖拽文件到此处上传<br>
                        <span>支持 .xlsx 格式</span>
                    </p>
                    <input type="file" name="file" id="fileInput" accept=".xlsx">
                    <p id="fileName" style="margin-top: 10px; color: #667eea; font-weight: 500;"></p>
                </div>
                <div class="submit-area">
                    <button type="submit" class="btn" id="submitBtn">开始分析</button>
                </div>
            </form>

            <!-- 进度显示区域 -->
            <div class="progress-container" id="progressContainer">
                <div class="progress-status" id="progressStatus">准备中...</div>
                <div class="progress-bar-bg">
                    <div class="progress-bar" id="progressBar"></div>
                </div>
                <div class="progress-percent" id="progressPercent">0%</div>
            </div>

            <!-- 结果显示区域 -->
            <div id="resultArea"></div>

            <div class="note">
                <h3>📋 使用说明</h3>
                <ul>
                    <li>上传包含 <strong>Media ID</strong>、<strong>Date</strong>、<strong>Source</strong>、<strong>Potential Audience</strong> 列的Excel文件</li>
                    <li>Media ID 对应品牌名称，Source 对应媒体名称</li>
                    <li>触达（人次）以 Date 最晚的数据为准</li>
                    <li>结果按声量（篇）降序排列，每个品牌显示 TOP10</li>
                    <li>触达单位：K（千）、M（百万），数值保留一位小数，不四舍五入</li>
                </ul>
            </div>
        </div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileName = document.getElementById('fileName');
        const submitBtn = document.getElementById('submitBtn');
        const progressContainer = document.getElementById('progressContainer');
        const progressStatus = document.getElementById('progressStatus');
        const progressBar = document.getElementById('progressBar');
        const progressPercent = document.getElementById('progressPercent');
        const resultArea = document.getElementById('resultArea');

        uploadArea.addEventListener('click', () => fileInput.click());

        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files.length) {
                fileInput.files = e.dataTransfer.files;
                fileName.textContent = e.dataTransfer.files[0].name;
            }
        });

        fileInput.addEventListener('change', () => {
            if (fileInput.files.length) {
                fileName.textContent = fileInput.files[0].name;
            }
        });

        document.getElementById('uploadForm').addEventListener('submit', async (e) => {
            e.preventDefault();

            if (!fileInput.files.length) {
                alert('请选择文件');
                return;
            }

            const formData = new FormData();
            formData.append('file', fileInput.files[0]);

            // 显示进度条
            submitBtn.disabled = true;
            submitBtn.textContent = '处理中，请稍候...';
            progressContainer.classList.add('active');
            progressBar.style.width = '50%';
            progressPercent.textContent = '50%';
            progressStatus.textContent = '正在分析数据...';
            resultArea.innerHTML = '';

            try {
                const response = await fetch('/analyze', {
                    method: 'POST',
                    body: formData
                });

                const html = await response.text();

                // 用返回的HTML替换整个页面
                document.open();
                document.write(html);
                document.close();

            } catch (error) {
                resultArea.innerHTML = '<div class="error">处理失败: ' + error.message + '</div>';
                submitBtn.disabled = false;
                submitBtn.textContent = '开始分析';
                progressContainer.classList.remove('active');
            }
        });
    </script>
</body>
</html>
'''


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/analyze', methods=['POST'])
def analyze():
    if 'file' not in request.files:
        return '''
            <div class="error">请选择文件</div>
            <script>setTimeout(() => { window.location.href = '/'; }, 2000)</script>
        '''

    file = request.files['file']
    if file.filename == '':
        return '''
            <div class="error">请选择文件</div>
            <script>setTimeout(() => { window.location.href = '/'; }, 2000)</script>
        '''

    if not file.filename.endswith('.xlsx'):
        return '''
            <div class="error">只支持 .xlsx 格式文件</div>
            <script>setTimeout(() => { window.location.href = '/'; }, 2000)</script>
        '''

    try:
        # 读取文件
        df = pd.read_excel(file)

        # 检查必要的列
        required_columns = ["Media ID", "Date", "Source", "Potential Audience"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            return '''
                <div class="error">缺少必要的列: ''' + ', '.join(missing_columns) + '''</div>
                <script>setTimeout(() => { window.location.href = '/'; }, 3000)</script>
            '''

        # 转换日期列并删除无效日期行
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date", "Media ID", "Source"])

        brands = df["Media ID"].unique().tolist()
        total_brands = len(brands)

        # 预计算最新触达数
        latest_audience_cache = {}
        for brand in brands:
            brand_df = df[df["Media ID"] == brand]
            for source in brand_df["Source"].unique():
                media_data = brand_df[brand_df["Source"] == source][["Date", "Potential Audience"]]
                if media_data.empty:
                    latest_audience_cache[(brand, source)] = None
                else:
                    latest_idx = media_data["Date"].idxmax()
                    latest_audience_cache[(brand, source)] = media_data.loc[latest_idx, "Potential Audience"]

        # 生成结果
        all_results = []
        for brand in brands:
            brand_df = df[df["Media ID"] == brand]
            volume_df = brand_df.groupby("Source").size().reset_index(name="声量（篇）")

            def get_latest_audience(media_name):
                return latest_audience_cache.get((brand, media_name), None)

            volume_df["触达（人次）"] = volume_df["Source"].apply(get_latest_audience)
            volume_df = volume_df.sort_values("声量（篇）", ascending=False).head(10)
            volume_df["品牌名称"] = brand
            volume_df["媒体类别"] = volume_df["Source"].apply(get_media_category)
            volume_df["触达（人次）"] = volume_df["触达（人次）"].apply(format_audience)
            volume_df.insert(0, "序号", range(1, len(volume_df) + 1))
            volume_df = volume_df.rename(columns={"Source": "媒体名称"})
            volume_df = volume_df[["序号", "品牌名称", "媒体名称", "声量（篇）", "触达（人次）", "媒体类别"]]
            all_results.append(volume_df)

        # 生成HTML
        brands_count = len(all_results)
        total_media = sum(len(r) for r in all_results)

        result_html = '''
            <div class="result-info">
                ✅ 分析完成！共处理 <strong>''' + str(brands_count) + '''</strong> 个品牌，生成 <strong>''' + str(total_media) + '''</strong> 条记录
            </div>
        '''

        for brand_df in all_results:
            brand_name = brand_df.iloc[0]['品牌名称']
            result_html += '''
                <div class="brand-section">
                    <div class="brand-title">''' + brand_name + ''' TOP10</div>
                    <table>
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>媒体名称</th>
                                <th>声量（篇）</th>
                                <th>触达（人次）</th>
                                <th>媒体类别</th>
                            </tr>
                        </thead>
                        <tbody>
            '''
            for _, row in brand_df.iterrows():
                category_class = '汽车' if '汽车' in row['媒体类别'] else ('商业' if '商业' in row['媒体类别'] else ('生活' if '生活' in row['媒体类别'] else '综合'))
                result_html += '''
                            <tr>
                                <td>''' + str(row['序号']) + '''</td>
                                <td>''' + str(row['媒体名称']) + '''</td>
                                <td>''' + str(row['声量（篇）']) + '''</td>
                                <td>''' + str(row['触达（人次）']) + '''</td>
                                <td><span class="category-tag category-''' + category_class + '''">''' + str(row['媒体类别']) + '''</span></td>
                            </tr>
                '''
            result_html += '''
                        </tbody>
                    </table>
                </div>
            '''

        # 返回完整页面
        return '''
            <div style="max-width: 900px; margin: 20px auto;">
                <div style="text-align: center; margin-bottom: 20px;">
                    <a href="/" style="display: inline-block; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 24px; border-radius: 8px; text-decoration: none; font-weight: 600;">← 上传新文件</a>
                </div>
                ''' + result_html + '''
            </div>
        '''

    except Exception as e:
        return '''
            <div class="error">处理失败: ''' + str(e) + '''</div>
            <script>setTimeout(() => { window.location.href = '/'; }, 5000)</script>
        '''


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
