import streamlit as st
import google.generativeai as genai
from PIL import Image
import io
import matplotlib.pyplot as plt
import matplotlib as mpl
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

# ==========================================
# 0. 全局配置与字体优化
# ==========================================
st.set_page_config(
    page_title="MathScan Pro - 印刷级公式还原",
    page_icon="✒️",
    layout="wide"
)

# 关键设置：让 Matplotlib 生成看起来像 LaTeX 标准排版的公式
# 使用 'cm' (Computer Modern) 字体，这是学术界标准数学字体
try:
    mpl.rcParams['mathtext.fontset'] = 'cm'
    mpl.rcParams['font.family'] = 'serif'
except:
    pass

# ==========================================
# 1. 核心 AI 识别模块
# ==========================================
def extract_content_with_gemini(api_key, image):
    """
    调用 Gemini 识别图片，并强制要求分离公式与文本
    """
    try:
        genai.configure(api_key=api_key)
        # 使用 Flash 模型，速度快且对文字识别极其精准
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = """
        你是一个专业的排版转换工具。请将图片内容转换为文档，要求如下：
        
        1. 【识别精准】：识别图片中的所有中文、英文文本和数学公式。
        2. 【LaTeX 格式】：所有数学公式必须写成标准的 LaTeX 格式。
           - 简单的变量（如 x, N, t）用单美元符号：$x$
           - 复杂的公式（如带上下标、求和、分数）用双美元符号：$$ ... $$
        3. 【保持结构】：严格保持原文的段落结构。
        4. 【输出纯净】：只输出内容，不要任何 Markdown 代码块标记（如 ```latex），不要任何开场白。
        """
        
        with st.spinner('🔍 AI 正在进行像素级识别与排版分析...'):
            response = model.generate_content([prompt, image])
            # 清理可能残留的 markdown 标记
            text = response.text.replace("```latex", "").replace("```markdown", "").replace("```", "")
            return text.strip()
    except Exception as e:
        st.error(f"连接 AI 服务失败: {str(e)}")
        return None

# ==========================================
# 2. 印刷级公式渲染模块
# ==========================================
def render_latex_high_quality(latex_str, dpi=400):
    """
    将 LaTeX 渲染为【透明背景】、【高清晰度】的图片流。
    视觉效果等同于印刷体，完全没有“照片感”。
    """
    try:
        # 1. 预处理：去掉可能导致渲染错误的命令
        latex_str = latex_str.strip()
        
        # 2. 创建画布 (尺寸会自动裁剪，初始值不重要)
        fig = plt.figure(figsize=(0.1, 0.1))
        
        # 3. 渲染配置
        # 加上 $ 符号启用数学模式
        # 使用 color='black' 确保字是纯黑
        text_content = f"${latex_str}$"
        
        # 4. 绘制文字
        # fontsize=18 保证清晰度足够高
        fig.text(0.5, 0.5, text_content, fontsize=18, ha='center', va='center', color='black')
        
        # 5. 关闭坐标轴
        plt.axis('off')
        
        # 6. 保存到内存
        buf = io.BytesIO()
        # transparent=True 关键：透明背景
        # bbox_inches='tight' 关键：紧贴公式边缘裁剪，不要留白
        plt.savefig(buf, format='png', dpi=dpi, bbox_inches='tight', pad_inches=0.05, transparent=True)
        plt.close(fig)
        buf.seek(0)
        return buf
    except Exception as e:
        # 如果某些复杂 LaTeX 无法被 matplotlib 渲染，返回 None
        print(f"渲染错误: {e}")
        plt.close(fig)
        return None

# ==========================================
# 3. 文档生成模块
# ==========================================
def create_professional_doc(raw_text):
    """
    生成 Word 文档，混合排版文字和高清公式
    """
    doc = Document()
    
    # 全局字体设置：中文微软雅黑，西文 Times New Roman
    style = doc.styles['Normal']
    style.font.name = 'Microsoft YaHei'
    style.font.size = Pt(11)
    
    # 拆解文本：找到 $$ 包裹的公式块
    # 正则逻辑：匹配 $$...$$ 之间的内容
    parts = re.split(r'(\$\$[\s\S]*?\$\$)', raw_text)
    
    for part in parts:
        part = part.strip()
        if not part:
            continue
            
        if part.startswith('$$') and part.endswith('$$'):
            # === 处理独立公式 ===
            latex_code = part[2:-2].strip() # 去掉 $$
            
            # 尝试渲染为高清图
            img_stream = render_latex_high_quality(latex_code)
            
            if img_stream:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(img_stream, height=Inches(0.6)) # 高度固定，看起来整齐
            else:
                # 渲染失败兜底：直接写 LaTeX 源码，标红
                p = doc.add_paragraph(latex_code)
                p.runs[0].font.color.rgb = RGBColor(200, 0, 0)
                
        else:
            # === 处理普通文本和行内公式 ===
            # 对于行内公式 $...$，为了保持段落连贯性，我们直接作为文本写入。
            # 这样 Word 里可以编辑，也不会因为插入小图片导致行间距错乱。
            doc.add_paragraph(part)
            
    # 保存
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ==========================================
# 4. 界面主逻辑
# ==========================================
st.title("✒️ MathScan Pro：印刷级公式还原")

# --- 修改重点：将 API Key 输入框移至主页面顶部，避免侧边栏报错遮挡 ---
with st.container():
    st.info("💡 请在下方输入密钥开始使用 (已从侧边栏移至此处，防止遮挡)")
    col1, col2 = st.columns([3, 1])
    with col1:
        api_key = st.text_input("🔑 Google API Key", type="password", help="在此粘贴 AIza 开头的密钥")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True) # 占位对齐
        st.markdown("[👉 免费获取 Key](https://aistudio.google.com/app/apikey)")

st.markdown("---")
st.markdown("上传图片，一键提取为 Word。公式将以**高清印刷格式**呈现（非截图）。")

# 主区域
uploaded_file = st.file_uploader("上传包含公式的图片 (截图/照片)", type=["png", "jpg", "jpeg"])

if uploaded_file:
    # 预览
    image = Image.open(uploaded_file)
    st.image(image, caption="原始图片", width=400)
    
    if api_key:
        if st.button("开始转换 (Generate Doc)", type="primary"):
            # 1. 识别
            text_result = extract_content_with_gemini(api_key, image)
            
            if text_result:
                st.success("✅ 识别完成！正在进行排版渲染...")
                
                # 2. 预览识别到的 LaTeX 源码
                with st.expander("查看 LaTeX 源码 (调试用)"):
                    st.code(text_result, language='latex')
                
                # 3. 生成文档
                doc_file = create_professional_doc(text_result)
                
                # 4. 下载
                st.download_button(
                    label="📥 下载最终 Word 文档",
                    data=doc_file,
                    file_name="MathScan_Result.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    else:
        st.warning("👈 请在上方输入框粘贴 API Key 才能开始工作")