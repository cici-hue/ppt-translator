import streamlit as st
from pptx import Presentation
import openai
import re
import tempfile
import os
from pptx.util import Pt
from dotenv import load_dotenv
import uuid

# ====================== 1. API密钥安全配置（不变，核心保留） ======================
if os.path.exists(".env"):
    load_dotenv()
DEEPSEEK_API_KEY = st.secrets.get("DEEPSEEK_API_KEY") or os.getenv("DEEPSEEK_API_KEY")
if not DEEPSEEK_API_KEY:
    st.error("❌ 未配置DeepSeek API密钥！请检查环境变量或Streamlit Secrets。")
    st.stop()

client = openai.OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

# ====================== 2. 【核心升级】多语言配置 + 目标语言-字体映射（必含德/泰/土耳其/孟加拉/越南语） ======================
# 语言配置：{显示名: (DeepSeek标准代码, 语言简称)} ｜ DeepSeek支持所有标准ISO 639-1代码
LANGUAGE_CONFIG = {
    "中文": ("zh", "Chinese"),
    "英语": ("en", "English"),
    "德语": ("de", "German"),       # 必含
    "泰语": ("th", "Thai"),         # 必含
    "土耳其语": ("tr", "Turkish"),   # 必含
    "孟加拉语": ("bn", "Bengali"),  # 必含
    "越南语": ("vi", "Vietnamese"), # 必含
    "法语": ("fr", "French"),
    "西班牙语": ("es", "Spanish"),
    "俄语": ("ru", "Russian"),
    "日语": ("ja", "Japanese"),
    "韩语": ("ko", "Korean")
}
# 目标语言-适配字体映射 ｜ 核心：系统原生字体，避免乱码，无需额外安装
# 西语/德语/土耳其语：Calibri（支持特殊字符）；亚洲语言：专属兼容字体
FONT_MAP = {
    "zh": "微软雅黑",       # 目标为中文
    "en": "Calibri",        # 目标为英语
    "de": "Calibri",        # 目标为德语
    "tr": "Calibri",        # 目标为土耳其语
    "fr": "Calibri",        # 目标为法语
    "es": "Calibri",        # 目标为西班牙语
    "ru": "Calibri",        # 目标为俄语
    "th": "TH Sarabun New", # 目标为泰语（Windows/macOS原生）
    "vi": "VN Times",       # 目标为越南语（Windows原生，macOS用Times New Roman兼容）
    "bn": "Siyam Rupali",   # 目标为孟加拉语（Windows/macOS原生）
    "ja": "MS Mincho",      # 目标为日语（Windows原生）
    "ko": "Malgun Gothic"   # 目标为韩语（Windows原生）
}
# 提取语言显示名（用于Streamlit下拉框）
LANG_NAMES = list(LANGUAGE_CONFIG.keys())

# ====================== 3. 工具函数（仅适配多语言，核心逻辑不变） ======================
def adjust_text_overflow_mild(text_frame, min_font_size=10):
    """温和溢出调整（不变）"""
    if not text_frame or not text_frame.text.strip():
        return
    text_frame.word_wrap = True
    src_sizes = [run.font.size for para in text_frame.paragraphs for run in para.runs if run.font.size is not None]
    if not src_sizes:
        return
    current_font = src_sizes[0]
    for _ in range(6):
        try:
            if text_frame.height >= text_frame.text_height:
                break
        except:
            break
        new_font = current_font - Pt(1)
        new_font = new_font if new_font >= Pt(min_font_size) else Pt(min_font_size)
        for para in text_frame.paragraphs:
            for run in para.runs:
                if run.font.size is not None:
                    run.font.size = new_font
        current_font = new_font
    if current_font == Pt(min_font_size):
        try:
            if text_frame.height < text_frame.text_height:
                st.warning(f"💡 部分文本略有溢出（已缩至最小10pt），建议手动微调文本框宽度")
        except:
            pass

def translate_text(text, src_lang_code, src_lang_name, tgt_lang_code, tgt_lang_name):
    """【多语言适配】翻译函数 | 传递语言代码+名称，去掉字符过滤（用户自主选择更精准）"""
    if not text or not text.strip():
        return text
    try:
        # 动态适配源/目标语言的翻译提示
        system_prompt = f"""你是专业的多语言翻译专家，精通{src_lang_name}和{tgt_lang_name}互译，严格遵循以下规则：
1. 术语准确：商务/办公PPT专业术语使用行业标准译法，保持一致性；
2. 格式保留：原文的换行、空格、标点、数字/单位完全不变，不增删任何内容；
3. 表达适配：符合目标语言的PPT阅读习惯，标题简洁有力，正文流畅自然；
4. 无额外输出：仅返回翻译结果，不添加解释、备注、标点修正等无关内容；
5. 特殊字符：准确处理目标语言的特殊字符/重音符号（如德语变音、越南语声调）。"""
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": text}],
            temperature=0.1,  # 低温度保证翻译结果稳定
            max_tokens=3000    # 增大token限制，适配多语言长文本
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"❌ 翻译出错：{str(e)}")
        return text

def translate_ppt(input_file_path, output_file_path, src_lang, tgt_lang):
    """【多语言核心】PPT翻译逻辑 | 解析源/目标语言的代码+名称，动态匹配字体"""
    # 解析源/目标语言的配置（代码+名称）
    src_lang_code, src_lang_name = LANGUAGE_CONFIG[src_lang]
    tgt_lang_code, tgt_lang_name = LANGUAGE_CONFIG[tgt_lang]
    # 动态匹配目标字体（解决多语言乱码）
    target_font = FONT_MAP[tgt_lang_code]
    
    try:
        prs = Presentation(input_file_path)
        st.success(f"✅ 成功加载PPT | 共{len(prs.slides)}张幻灯片 | 源语言：{src_lang} | 目标语言：{tgt_lang} | 适配字体：{target_font}")
    except Exception as e:
        st.error(f"❌ 加载PPT失败：{str(e)}")
        return False
    
    total_texts, translated_texts = 0, 0
    # 进度条+状态提示（不变，用户体验友好）
    progress_bar = st.progress(0)
    status_text = st.empty()

    for slide_idx, slide in enumerate(prs.slides, 1):
        status_text.text(f"🔄 处理第 {slide_idx}/{len(prs.slides)} 张幻灯片...")
        progress_bar.progress(slide_idx / len(prs.slides))

        for shape in slide.shapes:
            # 处理文本框（多语言字体适配，格式保留不变）
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    original_text = paragraph.text.strip()
                    if original_text:
                        total_texts += 1
                        # 调用多语言翻译函数
                        translated_text = translate_text(original_text, src_lang_code, src_lang_name, tgt_lang_code, tgt_lang_name)
                        if translated_text and translated_text != original_text:
                            # 保留原格式（加粗/颜色/字号，不变）
                            src_font = paragraph.runs[0].font if paragraph.runs else None
                            paragraph.text = ""
                            new_run = paragraph.add_run()
                            new_run.text = translated_text
                            if src_font:
                                new_run.font.bold = src_font.bold if src_font.bold is not None else False
                                new_run.font.size = src_font.size
                                new_run.font.name = target_font  # 多语言字体适配
                                try:
                                    new_run.font.color.rgb = src_font.color.rgb
                                except:
                                    pass
                            # 1倍行间距（不变，全版本支持）
                            paragraph.line_spacing = 1
                            # 温和溢出调整（不变）
                            adjust_text_overflow_mild(shape.text_frame)
                            translated_texts += 1
            # 处理表格（与文本框完全一致，多语言字体+格式保留）
            if shape.has_table:
                try:
                    table = shape.table
                    for row_idx, row in enumerate(table.rows):
                        for cell_idx, cell in enumerate(row.cells):
                            cell_text = cell.text.strip()
                            if cell_text:
                                total_texts += 1
                                translated_cell = translate_text(cell_text, src_lang_code, src_lang_name, tgt_lang_code, tgt_lang_name)
                                if translated_cell and translated_cell != cell_text:
                                    cell_src_font = None
                                    if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
                                        cell_src_font = cell.text_frame.paragraphs[0].runs[0].font
                                    cell.text_frame.clear()
                                    cell_para = cell.text_frame.add_paragraph()
                                    cell_run = cell_para.add_run()
                                    cell_run.text = translated_cell
                                    # 格式保留+多语言字体
                                    if cell_src_font:
                                        cell_run.font.bold = cell_src_font.bold if cell_src_font.bold is not None else False
                                        cell_run.font.size = cell_src_font.size
                                        cell_run.font.name = target_font
                                        try:
                                            cell_run.font.color.rgb = cell_src_font.color.rgb
                                        except:
                                            pass
                                    cell_para.line_spacing = 1
                                    adjust_text_overflow_mild(cell.text_frame)
                                    translated_texts += 1
                except Exception as e:
                    st.warning(f"⚠️ 表格处理异常（跳过）：{str(e)[:40]}...")

    # 保存翻译后的PPT（不变）
    try:
        prs.save(output_file_path)
        progress_bar.progress(100)
        status_text.text("✅ 翻译完成！")
        # 多语言翻译统计（动态显示）
        st.success(f"""
        📊 翻译统计结果 | 源语言：{src_lang} → 目标语言：{tgt_lang}
        ├─ 总文本块（文本框+表格）：{total_texts}
        ├─ 成功翻译文本块：{translated_texts}
        ├─ 目标语言适配字体：{target_font}
        └─ 格式保留：加粗/颜色/字号1:1保留 + 1倍行间距 + 温和溢出调整
        """)
        return True
    except Exception as e:
        st.error(f"❌ 保存PPT失败：{str(e)}（请关闭本地同名PPT文件后重试）")
        return False

# ====================== 4. Streamlit Web交互界面（多语言下拉框，操作不变） ======================
def main():
    st.set_page_config(page_title="PPT智能翻译工具", page_icon="📄", layout="wide")
    st.title("📄 PPT智能翻译工具")
    st.divider()

    # 侧边栏：【多语言升级】源/目标语言下拉选择框 + 功能说明
    with st.sidebar:
        st.header("⚙️ 翻译配置")
        # 多语言源语言选择（默认中文）
        src_lang = st.selectbox("🔤 源语言", LANG_NAMES, index=LANG_NAMES.index("中文"))
        # 多语言目标语言选择（默认英语）
        tgt_lang = st.selectbox("🌐 目标语言", LANG_NAMES, index=LANG_NAMES.index("英语"))
        # 校验：源语言≠目标语言
        if src_lang == tgt_lang:
            st.error("❌ 源语言和目标语言不能相同，请重新选择！")
            st.stop()
        # 功能说明（适配多语言）
        st.info("""
        📌 核心功能说明
        1. 支持12种主流语言互译；
        2. 自动适配目标语言字体，避免乱码；
        3. 保留原PPT所有格式；
        4. 支持文本框/表格翻译；
        5. 仅支持.pptx格式，文件上传后一键翻译、下载结果。
        """)
        st.warning("""
        ⚠️ 温馨提示
        1. 建议上传小于20MB的PPT文件，翻译速度更快；
        2. 复杂艺术字/特殊形状文本可能无法解析（属python-pptx库限制）；
        3. 翻译结果请自行核对专业术语，确保准确性。
        """)

    # 主界面：文件上传（不变，仅支持.pptx）
    st.subheader("📤 上传PPT文件（仅支持.pptx格式）")
    uploaded_file = st.file_uploader("点击选择或拖拽PPT文件至此处", type=["pptx"], accept_multiple_files=False)

    if uploaded_file is not None:
        # 显示上传文件信息
        file_size = round(uploaded_file.size / 1024 / 1024, 2)
        st.info(f"📁 已上传文件：{uploaded_file.name} | 文件大小：{file_size} MB")
        # 生成唯一临时文件名（避免冲突，不变）
        unique_id = str(uuid.uuid4())[:8]
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_input:
            temp_input.write(uploaded_file.getbuffer())
            temp_input_path = temp_input.name

        # 翻译按钮（主按钮，醒目）
        if st.button("🚀 开始多语言翻译", type="primary", use_container_width=True):
            # 生成输出临时文件
            temp_output_path = os.path.join(tempfile.gettempdir(), f"ppt_translated_{unique_id}.pptx")
            # 执行多语言翻译
            translate_success = translate_ppt(temp_input_path, temp_output_path, src_lang, tgt_lang)
            # 提供下载链接（动态生成文件名，如"原文件名_中译德.pptx"）
            if translate_success and os.path.exists(temp_output_path):
                download_file_name = f"{os.path.splitext(uploaded_file.name)[0]}_{src_lang}译{tgt_lang}.pptx"
                with open(temp_output_path, "rb") as f:
                    st.download_button(
                        label="📥 下载翻译后的PPT文件",
                        data=f,
                        file_name=download_file_name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True
                    )
            # 清理临时文件（避免占用磁盘，不变）
            os.unlink(temp_input_path)
            if os.path.exists(temp_output_path):
                os.unlink(temp_output_path)

if __name__ == "__main__":
    main()
