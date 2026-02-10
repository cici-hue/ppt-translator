# import streamlit as st
# from pptx import Presentation
# import openai
# import re
# import tempfile
# import os
# from pptx.util import Pt
# from dotenv import load_dotenv
# import uuid

# # ====================== 1. API Key Security Configuration (Unchanged) ======================
# if os.path.exists(".env"):
#     load_dotenv()
# DEEPSEEK_API_KEY = st.secrets.get("DEEPSEEK_API_KEY") or os.getenv("DEEPSEEK_API_KEY")
# if not DEEPSEEK_API_KEY:
#     st.error("❌ DeepSeek API Key not configured! Please check environment variables or Streamlit Secrets.")
#     st.stop()

# client = openai.OpenAI(
#     api_key=DEEPSEEK_API_KEY,
#     base_url="https://api.deepseek.com"
# )

# # ====================== 2. Multi-Language Config + Font Mapping (Updated for English UI) ======================
# # Language Config: {Display Name (English): (DeepSeek Standard Code, Language English Name)}
# # Required languages: Chinese/English/German/Thai/Turkish/Bengali/Vietnamese
# LANGUAGE_CONFIG = {
#     "Chinese": ("zh", "Chinese"),
#     "English": ("en", "English"),
#     "German": ("de", "German"),       # Required
#     "Thai": ("th", "Thai"),           # Required
#     "Turkish": ("tr", "Turkish"),     # Required
#     "Bengali": ("bn", "Bengali"),     # Required
#     "Vietnamese": ("vi", "Vietnamese"), # Required
#     "French": ("fr", "French"),
#     "Spanish": ("es", "Spanish"),
#     "Russian": ("ru", "Russian"),
#     "Japanese": ("ja", "Japanese"),
#     "Korean": ("ko", "Korean")
# }
# # Target Language - Font Mapping | Core: System-native fonts to avoid garbled text
# FONT_MAP = {
#     "zh": "微软雅黑",       # Target: Chinese
#     "en": "Calibri",        # Target: English
#     "de": "Calibri",        # Target: German
#     "tr": "Calibri",        # Target: Turkish
#     "fr": "Calibri",        # Target: French
#     "es": "Calibri",        # Target: Spanish
#     "ru": "Calibri",        # Target: Russian
#     "th": "TH Sarabun New", # Target: Thai (Native to Windows/macOS)
#     "vi": "VN Times",       # Target: Vietnamese (Native to Windows, compatible with Times New Roman on macOS)
#     "bn": "Siyam Rupali",   # Target: Bengali (Native to Windows/macOS)
#     "ja": "MS Mincho",      # Target: Japanese (Native to Windows)
#     "ko": "Malgun Gothic"   # Target: Korean (Native to Windows)
# }
# # Extract language display names (for Streamlit dropdown)
# LANG_NAMES = list(LANGUAGE_CONFIG.keys())

# # ====================== 3. Utility Functions (Minimal English Adaptation) ======================
# def adjust_text_overflow_mild(text_frame, min_font_size=10):
#     """Mild text overflow adjustment (Unchanged)"""
#     if not text_frame or not text_frame.text.strip():
#         return
#     text_frame.word_wrap = True
#     src_sizes = [run.font.size for para in text_frame.paragraphs for run in para.runs if run.font.size is not None]
#     if not src_sizes:
#         return
#     current_font = src_sizes[0]
#     for _ in range(6):
#         try:
#             if text_frame.height >= text_frame.text_height:
#                 break
#         except:
#             break
#         new_font = current_font - Pt(1)
#         new_font = new_font if new_font >= Pt(min_font_size) else Pt(min_font_size)
#         for para in text_frame.paragraphs:
#             for run in para.runs:
#                 if run.font.size is not None:
#                     run.font.size = new_font
#         current_font = new_font
#     if current_font == Pt(min_font_size):
#         try:
#             if text_frame.height < text_frame.text_height:
#                 st.warning(f"💡 Some text is slightly overflowing (reduced to minimum 10pt). Please manually adjust text box width.")
#         except:
#             pass

# def translate_text(text, src_lang_code, src_lang_name, tgt_lang_code, tgt_lang_name):
#     """Multi-language translation function | Pass language code + name (English UI adaptation)"""
#     if not text or not text.strip():
#         return text
#     try:
#         # Dynamically adapt translation prompt to source/target language
#         system_prompt = f"""You are a professional multilingual translation expert, proficient in {src_lang_name} and {tgt_lang_name} mutual translation. Strictly follow these rules:
# 1. Terminology accuracy: Use industry-standard translations for business/office PPT professional terms, maintain consistency;
# 2. Format preservation: Keep line breaks, spaces, punctuation, numbers/units in the original text unchanged, no additions or deletions;
# 3. Expression adaptation: Conform to PPT reading habits of the target language, concise and powerful titles, fluent text;
# 4. No extra output: Only return translation results, no explanations, notes, punctuation corrections or irrelevant content;
# 5. Special characters: Accurately handle special characters/diacritics of the target language (e.g., German umlauts, Vietnamese tones)."""
#         response = client.chat.completions.create(
#             model="deepseek-chat",
#             messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": text}],
#             temperature=0.1,  # Low temperature ensures stable translation results
#             max_tokens=3000    # Increase token limit for multi-language long text
#         )
#         return response.choices[0].message.content.strip()
#     except Exception as e:
#         st.error(f"❌ Translation error: {str(e)}")
#         return text

# def translate_ppt(input_file_path, output_file_path, src_lang, tgt_lang):
#     """Core PPT translation logic | Parse source/target language config, dynamic font matching"""
#     # Parse source/target language config (code + name)
#     src_lang_code, src_lang_name = LANGUAGE_CONFIG[src_lang]
#     tgt_lang_code, tgt_lang_name = LANGUAGE_CONFIG[tgt_lang]
#     # Dynamically match target font (solve garbled text)
#     target_font = FONT_MAP[tgt_lang_code]
    
#     try:
#         prs = Presentation(input_file_path)
#         st.success(f"✅ PPT loaded successfully | Total slides: {len(prs.slides)} | Source Language: {src_lang} | Target Language: {tgt_lang} | Adapted Font: {target_font}")
#     except Exception as e:
#         st.error(f"❌ Failed to load PPT: {str(e)}")
#         return False
    
#     total_texts, translated_texts = 0, 0
#     # Progress bar + status prompt (English adaptation)
#     progress_bar = st.progress(0)
#     status_text = st.empty()

#     for slide_idx, slide in enumerate(prs.slides, 1):
#         status_text.text(f"🔄 Processing slide {slide_idx}/{len(prs.slides)}...")
#         progress_bar.progress(slide_idx / len(prs.slides))

#         for shape in slide.shapes:
#             # Process text boxes (multi-language font adaptation, format preservation)
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     original_text = paragraph.text.strip()
#                     if original_text:
#                         total_texts += 1
#                         # Call multi-language translation function
#                         translated_text = translate_text(original_text, src_lang_code, src_lang_name, tgt_lang_code, tgt_lang_name)
#                         if translated_text and translated_text != original_text:
#                             # Preserve original format (bold/color/font size)
#                             src_font = paragraph.runs[0].font if paragraph.runs else None
#                             paragraph.text = ""
#                             new_run = paragraph.add_run()
#                             new_run.text = translated_text
#                             if src_font:
#                                 new_run.font.bold = src_font.bold if src_font.bold is not None else False
#                                 new_run.font.size = src_font.size
#                                 new_run.font.name = target_font  # Multi-language font adaptation
#                                 try:
#                                     new_run.font.color.rgb = src_font.color.rgb
#                                 except:
#                                     pass
#                             # 1x line spacing (supported in all versions)
#                             paragraph.line_spacing = 1
#                             # Mild overflow adjustment
#                             adjust_text_overflow_mild(shape.text_frame)
#                             translated_texts += 1
#             # Process tables (same as text boxes, multi-language font + format preservation)
#             if shape.has_table:
#                 try:
#                     table = shape.table
#                     for row_idx, row in enumerate(table.rows):
#                         for cell_idx, cell in enumerate(row.cells):
#                             cell_text = cell.text.strip()
#                             if cell_text:
#                                 total_texts += 1
#                                 translated_cell = translate_text(cell_text, src_lang_code, src_lang_name, tgt_lang_code, tgt_lang_name)
#                                 if translated_cell and translated_cell != cell_text:
#                                     cell_src_font = None
#                                     if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
#                                         cell_src_font = cell.text_frame.paragraphs[0].runs[0].font
#                                     cell.text_frame.clear()
#                                     cell_para = cell.text_frame.add_paragraph()
#                                     cell_run = cell_para.add_run()
#                                     cell_run.text = translated_cell
#                                     # Format preservation + multi-language font
#                                     if cell_src_font:
#                                         cell_run.font.bold = cell_src_font.bold if cell_src_font.bold is not None else False
#                                         cell_run.font.size = cell_src_font.size
#                                         cell_run.font.name = target_font
#                                         try:
#                                             cell_run.font.color.rgb = cell_src_font.color.rgb
#                                         except:
#                                             pass
#                                     cell_para.line_spacing = 1
#                                     adjust_text_overflow_mild(cell.text_frame)
#                                     translated_texts += 1
#                 except Exception as e:
#                     st.warning(f"⚠️ Table processing exception (skipped): {str(e)[:40]}...")

#     # Save translated PPT (English adaptation)
#     try:
#         prs.save(output_file_path)
#         progress_bar.progress(100)
#         status_text.text("✅ Translation completed!")
#         # Multi-language translation statistics (English format)
#         st.success(f"""
#         📊 Translation Statistics | Source Language: {src_lang} → Target Language: {tgt_lang}
#         ├─ Total text blocks (text boxes + tables): {total_texts}
#         ├─ Successfully translated text blocks: {translated_texts}
#         ├─ Target language adapted font: {target_font}
#         └─ Format preservation: 1:1 retention of bold/color/font size + 1x line spacing + mild overflow adjustment
#         """)
#         return True
#     except Exception as e:
#         st.error(f"❌ Failed to save PPT: {str(e)} (Please close the target PPT file and try again)")
#         return False

# # ====================== 4. Streamlit Web Interface (Fully English) ======================
# def main():
#     st.set_page_config(page_title="PPT Smart Translation Tool", page_icon="📄", layout="wide")
#     st.title("📄 PPT Smart Translation Tool")
#     st.divider()

#     # Sidebar: Translation Configuration (Fully English)
#     with st.sidebar:
#         st.header("⚙️ Translation Settings")
#         # Multi-language source language selection (default: Chinese)
#         src_lang = st.selectbox("🔤 Source Language", LANG_NAMES, index=LANG_NAMES.index("Chinese"))
#         # Multi-language target language selection (default: English)
#         tgt_lang = st.selectbox("🌐 Target Language", LANG_NAMES, index=LANG_NAMES.index("English"))
#         # Validation: Source language != Target language
#         if src_lang == tgt_lang:
#             st.error("❌ Source language and target language cannot be the same! Please reselect.")
#             st.stop()
#         # Feature description (English adaptation)
#         st.info("""
#         📌 Core Features
#         1. Supports mutual translation between 12 mainstream languages;
#         2. Automatically adapts target language fonts to avoid garbled text;
#         3. Preserves all original PPT formats;
#         4. Supports text box/table translation;
#         5. Only supports .pptx format - upload file, translate with one click, download results.
#         """)
#         st.warning("""
#         ⚠️ Important Notes
#         1. Recommended to upload PPT files smaller than 20MB for faster translation;
#         2. Complex artistic text/special shape text may not be parsed (python-pptx library limitation);
#         3. Please verify professional terminology in translation results to ensure accuracy.
#         """)

#     # Main Interface: File Upload (Fully English)
#     st.subheader("📤 Upload PPT File (Only .pptx format supported)")
#     uploaded_file = st.file_uploader("Click to select or drag PPT file here", type=["pptx"], accept_multiple_files=False)

#     if uploaded_file is not None:
#         # Display uploaded file information (English)
#         file_size = round(uploaded_file.size / 1024 / 1024, 2)
#         st.info(f"📁 Uploaded File: {uploaded_file.name} | File Size: {file_size} MB")
#         # Generate unique temporary file name (avoid conflicts)
#         unique_id = str(uuid.uuid4())[:8]
#         with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_input:
#             temp_input.write(uploaded_file.getbuffer())
#             temp_input_path = temp_input.name

#         # Translation button (prominent, English)
#         if st.button("🚀 Start Multi-Language Translation", type="primary", use_container_width=True):
#             # Generate output temporary file
#             temp_output_path = os.path.join(tempfile.gettempdir(), f"ppt_translated_{unique_id}.pptx")
#             # Execute multi-language translation
#             translate_success = translate_ppt(temp_input_path, temp_output_path, src_lang, tgt_lang)
#             # Provide download link (dynamically generated file name, e.g., "filename_Chinese_to_German.pptx")
#             if translate_success and os.path.exists(temp_output_path):
#                 download_file_name = f"{os.path.splitext(uploaded_file.name)[0]}_{src_lang}_to_{tgt_lang}.pptx"
#                 with open(temp_output_path, "rb") as f:
#                     st.download_button(
#                         label="📥 Download Translated PPT File",
#                         data=f,
#                         file_name=download_file_name,
#                         mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
#                         type="primary",
#                         use_container_width=True
#                     )
#             # Clean up temporary files (avoid disk usage)
#             os.unlink(temp_input_path)
#             if os.path.exists(temp_output_path):
#                 os.unlink(temp_output_path)

# if __name__ == "__main__":
#     main()




import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import openai
import re
import tempfile
import os
from pptx.util import Pt
from dotenv import load_dotenv
import uuid

# ====================== 1. API Key Security Configuration ======================
if os.path.exists(".env"):
    load_dotenv()
DEEPSEEK_API_KEY = st.secrets.get("DEEPSEEK_API_KEY") or os.getenv("DEEPSEEK_API_KEY")
if not DEEPSEEK_API_KEY:
    st.error("❌ DeepSeek API Key not configured! Please check environment variables or Streamlit Secrets.")
    st.stop()

client = openai.OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

# ====================== 2. Multi-Language Config + Font Mapping ======================
LANGUAGE_CONFIG = {
    "Chinese": ("zh", "Chinese"),
    "English": ("en", "English"),
    "German": ("de", "German"),
    "Thai": ("th", "Thai"),
    "Turkish": ("tr", "Turkish"),
    "Bengali": ("bn", "Bengali"),
    "Vietnamese": ("vi", "Vietnamese"),
    "French": ("fr", "French"),
    "Spanish": ("es", "Spanish"),
    "Russian": ("ru", "Russian"),
    "Japanese": ("ja", "Japanese"),
    "Korean": ("ko", "Korean")
}

FONT_MAP = {
    "zh": "微软雅黑",
    "en": "Calibri",
    "de": "Calibri",
    "tr": "Calibri",
    "fr": "Calibri",
    "es": "Calibri",
    "ru": "Calibri",
    "th": "TH Sarabun New",
    "vi": "VN Times",
    "bn": "Siyam Rupali",
    "ja": "MS Mincho",
    "ko": "Malgun Gothic"
}

LANG_NAMES = list(LANGUAGE_CONFIG.keys())

# ====================== 3. Enhanced Utility Functions ======================
def adjust_text_overflow_mild(text_frame, min_font_size=10):
    """Mild text overflow adjustment"""
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
                st.warning(f"💡 Some text is slightly overflowing (reduced to minimum 10pt). Please manually adjust text box width.")
        except:
            pass

def translate_text(text, src_lang_code, src_lang_name, tgt_lang_code, tgt_lang_name):
    """Multi-language translation function"""
    if not text or not text.strip():
        return text
    try:
        system_prompt = f"""You are a professional multilingual translation expert, proficient in {src_lang_name} and {tgt_lang_name} mutual translation. Strictly follow these rules:
1. Terminology accuracy: Use industry-standard translations for business/office PPT professional terms, maintain consistency;
2. Format preservation: Keep line breaks, spaces, punctuation, numbers/units in the original text unchanged, no additions or deletions;
3. Expression adaptation: Conform to PPT reading habits of the target language, concise and powerful titles, fluent text;
4. No extra output: Only return translation results, no explanations, notes, punctuation corrections or irrelevant content;
5. Special characters: Accurately handle special characters/diacritics of the target language (e.g., German umlauts, Vietnamese tones)."""
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": text}],
            temperature=0.1,
            max_tokens=3000
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"❌ Translation error: {str(e)}")
        return text

def process_text_frame(text_frame, translate_func, target_font, stats):
    """统一处理文本框（普通文本/艺术字/SmartArt节点）"""
    if not text_frame or not text_frame.text.strip():
        return
    
    original_text = text_frame.text.strip()
    stats['total'] += 1
    
    # 执行翻译
    translated_text = translate_func(original_text)
    if not translated_text or translated_text == original_text:
        return
    
    # 保留格式并替换文本
    for para in text_frame.paragraphs:
        if para.runs:
            src_font = para.runs[0].font
            para.text = ""
            new_run = para.add_run()
            new_run.text = translated_text
            
            # 格式保留 + 目标字体适配
            new_run.font.bold = src_font.bold if src_font.bold is not None else False
            new_run.font.size = src_font.size
            new_run.font.name = target_font
            try:
                new_run.font.color.rgb = src_font.color.rgb
            except:
                pass
            para.line_spacing = 1
    
    adjust_text_overflow_mild(text_frame)
    stats['translated'] += 1

def process_smartart(smartart, translate_func, target_font, stats):
    """处理SmartArt图形文本"""
    try:
        # 遍历SmartArt所有节点
        for node in smartart.all_nodes:
            if node.text_frame and node.text_frame.text.strip():
                process_text_frame(node.text_frame, translate_func, target_font, stats)
    except Exception as e:
        st.warning(f"⚠️ SmartArt processing error: {str(e)[:50]}...")

def process_shape_recursive(shape, translate_func, target_font, stats):
    """递归处理所有形状（包括分组、艺术字、SmartArt）"""
    try:
        # 1. 处理分组形状（递归）
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in shape.shapes:
                process_shape_recursive(sub_shape, translate_func, target_font, stats)
        
        # 2. 处理SmartArt
        elif shape.shape_type == MSO_SHAPE_TYPE.SMARTART:
            process_smartart(shape.smartart, translate_func, target_font, stats)
        
        # 3. 处理艺术字/普通文本框（覆盖WordArt）
        elif shape.has_text_frame:
            process_text_frame(shape.text_frame, translate_func, target_font, stats)
        
        # 4. 处理表格
        elif shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    if cell.text_frame and cell.text_frame.text.strip():
                        process_text_frame(cell.text_frame, translate_func, target_font, stats)
    
    except Exception as e:
        st.warning(f"⚠️ Shape processing error ({shape.name}): {str(e)[:50]}...")

def translate_ppt(input_file_path, output_file_path, src_lang, tgt_lang):
    """Enhanced PPT translation logic with art text/SmartArt support"""
    src_lang_code, src_lang_name = LANGUAGE_CONFIG[src_lang]
    tgt_lang_code, tgt_lang_name = LANGUAGE_CONFIG[tgt_lang]
    target_font = FONT_MAP[tgt_lang_code]
    
    try:
        prs = Presentation(input_file_path)
        st.success(f"✅ PPT loaded successfully | Total slides: {len(prs.slides)} | Source: {src_lang} | Target: {tgt_lang} | Font: {target_font}")
    except Exception as e:
        st.error(f"❌ Failed to load PPT: {str(e)}")
        return False
    
    # 初始化统计信息
    stats = {'total': 0, 'translated': 0}
    progress_bar = st.progress(0)
    status_text = st.empty()

    # 定义翻译函数（闭包简化参数传递）
    def translate_wrapper(text):
        return translate_text(text, src_lang_code, src_lang_name, tgt_lang_code, tgt_lang_name)

    # 遍历所有幻灯片并递归处理形状
    for slide_idx, slide in enumerate(prs.slides, 1):
        status_text.text(f"🔄 Processing slide {slide_idx}/{len(prs.slides)}...")
        progress_bar.progress(slide_idx / len(prs.slides))
        
        # 递归处理当前幻灯片的所有形状
        for shape in slide.shapes:
            process_shape_recursive(shape, translate_wrapper, target_font, stats)

    # 保存翻译后的PPT
    try:
        prs.save(output_file_path)
        progress_bar.progress(100)
        status_text.text("✅ Translation completed!")
        
        st.success(f"""
        📊 Translation Statistics | {src_lang} → {tgt_lang}
        ├─ Total text blocks (text/table/art/SmartArt): {stats['total']}
        ├─ Successfully translated: {stats['translated']}
        ├─ Target font: {target_font}
        └─ Supported elements: text boxes, tables, WordArt, SmartArt, grouped shapes
        """)
        return True
    except Exception as e:
        st.error(f"❌ Failed to save PPT: {str(e)} (Please close the target PPT file and try again)")
        return False

# ====================== 4. Streamlit Web Interface ======================
def main():
    st.set_page_config(page_title="PPT Smart Translation Tool", page_icon="📄", layout="wide")
    st.title("📄 PPT Smart Translation Tool (Enhanced)")
    st.divider()

    with st.sidebar:
        st.header("⚙️ Translation Settings")
        src_lang = st.selectbox("🔤 Source Language", LANG_NAMES, index=LANG_NAMES.index("Chinese"))
        tgt_lang = st.selectbox("🌐 Target Language", LANG_NAMES, index=LANG_NAMES.index("English"))
        
        if src_lang == tgt_lang:
            st.error("❌ Source language and target language cannot be the same! Please reselect.")
            st.stop()
        
        st.info("""
        📌 Enhanced Features
        1. Supports 12 languages mutual translation;
        2. Translates WordArt/SmartArt/grouped shapes text;
        3. Preserves original formats & adapts fonts;
        4. Supports text boxes/tables/SmartArt/WordArt;
        5. Only .pptx format supported.
        """)
        st.warning("""
        ⚠️ Notes
        1. Files <20MB recommended for faster translation;
        2. Extremely complex custom shapes may not be parsed;
        3. Verify professional terminology in results.
        """)

    st.subheader("📤 Upload PPT File (Only .pptx format supported)")
    uploaded_file = st.file_uploader("Click to select or drag PPT file here", type=["pptx"], accept_multiple_files=False)

    if uploaded_file is not None:
        file_size = round(uploaded_file.size / 1024 / 1024, 2)
        st.info(f"📁 Uploaded File: {uploaded_file.name} | Size: {file_size} MB")
        
        unique_id = str(uuid.uuid4())[:8]
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_input:
            temp_input.write(uploaded_file.getbuffer())
            temp_input_path = temp_input.name

        if st.button("🚀 Start Multi-Language Translation", type="primary", use_container_width=True):
            temp_output_path = os.path.join(tempfile.gettempdir(), f"ppt_translated_{unique_id}.pptx")
            translate_success = translate_ppt(temp_input_path, temp_output_path, src_lang, tgt_lang)
            
            if translate_success and os.path.exists(temp_output_path):
                download_file_name = f"{os.path.splitext(uploaded_file.name)[0]}_{src_lang}_to_{tgt_lang}.pptx"
                with open(temp_output_path, "rb") as f:
                    st.download_button(
                        label="📥 Download Translated PPT File",
                        data=f,
                        file_name=download_file_name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True
                    )
            
            # 清理临时文件
            os.unlink(temp_input_path)
            if os.path.exists(temp_output_path):
                os.unlink(temp_output_path)

if __name__ == "__main__":
    main()
