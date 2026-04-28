import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
import io, base64, json, re, requests, uuid, copy, html
from datetime import datetime
from docx import Document
from pypdf import PdfReader
from streamlit_local_storage import LocalStorage

# ==========================================
# 1. 页面全局配置与前端美化
# ==========================================
st.set_page_config(page_title="ZenMux 创作者工作站", page_icon="🐙", layout="wide")
st.markdown("""
    <style>
    .stButton>button { border-radius: 8px; font-weight: bold; transition: all 0.3s; }
    .stChatInput { padding-bottom: 20px; }
    button[title="View fullscreen"] {display: none;}
    .css-1jc7ptx, .e1ewe7hr3, .viewerBadge_container__1QSob, .styles_viewerBadge__1yB5_ {display: none;}
    @media (max-width: 768px) {
        .block-container { padding-top: 2rem; padding-bottom: 5rem; }
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 数据持久化层 (LocalStorage 桥接) - 修复版
# ==========================================
localS = LocalStorage()

def trigger_save():
    st.session_state._needs_save = True

def execute_save():
    if st.session_state.get("_needs_save", False) and not st.session_state.get("is_streaming", False):
        try:
            data = {
                "profiles": st.session_state.profiles,
                "free_chats": st.session_state.free_chats
            }
            localS.setItem("zenmux_data", json.dumps(data, ensure_ascii=False))
            st.session_state._needs_save = False
        except Exception as e:
            pass  # 静默失败，避免打断流程

# 初始化状态
if "initialized" not in st.session_state:
    st.session_state.initialized = False
    st.session_state.ls_loaded = False
    st.session_state._needs_save = False
    st.session_state.is_streaming = False
    st.session_state.stop_streaming = False
    st.session_state._ls_retry_count = 0

# 🔥 修复：水合逻辑（带重试+超时兜底）
if not st.session_state.ls_loaded:
    # 默认数据
    default_profiles = [{
        "name": "默认引擎", "base_url": "", "api_key": "", "model": "anthropic/claude-3-5-sonnet-20240620",
        "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 8192,
        "use_top_p": False, "top_p": 1.0, "use_frequency_penalty": False, "frequency_penalty": 0.0
    }]
    first_id = str(uuid.uuid4())
    default_chats = {first_id: {"title": "新对话", "messages": [], "session_knowledge": [], "system_prompt": "", "is_pinned": False, "is_archived": False}}

    # 尝试读取本地存储
    saved_data = None
    try:
        saved_data = localS.getItem("zenmux_data")
    except Exception:
        saved_data = None

    st.session_state._ls_retry_count += 1

    # 🔥 核心修复：如果重试超过3次仍未ready，直接用默认值启动（避免死等）
    if saved_data is None and st.session_state._ls_retry_count < 3:
        # 组件可能还没ready，显示提示并重试
        placeholder = st.empty()
        placeholder.info(f"🔄 正在初始化本地存储... ({st.session_state._ls_retry_count}/3)")
        import time
        time.sleep(0.5)
        st.rerun()

    # 解析数据（即使重试超限也继续，用默认值）
    if saved_data and saved_data != "null":
        try:
            data = json.loads(saved_data) if isinstance(saved_data, str) else saved_data
            if isinstance(data, dict):
                st.session_state.profiles = data.get("profiles", default_profiles)
                st.session_state.free_chats = data.get("free_chats", default_chats)
            else:
                st.session_state.profiles = default_profiles
                st.session_state.free_chats = default_chats
        except Exception as e:
            st.session_state.profiles = default_profiles
            st.session_state.free_chats = default_chats
    else:
        st.session_state.profiles = default_profiles
        st.session_state.free_chats = default_chats

    # 确保至少有一个chat
    if not st.session_state.free_chats:
        st.session_state.free_chats = default_chats

    st.session_state.active_profile_idx = 0
    st.session_state.current_chat_id = list(st.session_state.free_chats.keys())[-1]
    st.session_state.current_page = "💬 自由聊天区"
    st.session_state.ls_loaded = True
    st.session_state.initialized = True
    st.rerun()

# ==========================================
# 3. 核心底层辅助函数
# ==========================================
def render_copy_button(text):
    b64 = base64.b64encode(text.encode("utf-8")).decode("utf-8")
    uid = uuid.uuid4().hex[:8]
    html_code = (
        '<div style="display:flex;justify-content:flex-end;padding-right:10px;">'
        '<button id="cb' + uid + '" onclick="(function(b){navigator.clipboard.writeText('
        "decodeURIComponent(escape(atob('" + b64 + "')))).then(function(){"
        "b.innerText='\\u2705 \\u5df2\\u590d\\u5236';b.style.color='#4CAF50';"
        "setTimeout(function(){b.innerText='\\ud83d\\udccb \\u590d\\u5236';b.style.color='#aaa';},2000);"
        '})})(this)" '
        'style="border:none;background:transparent;color:#aaa;cursor:pointer;font-size:12px;'
        'font-weight:bold;padding:5px 10px;border-radius:6px;">📋 复制</button></div>'
    )
    components.html(html_code, height=30)

def clean_novel_text(text):
    text = re.sub(r'^\s*(好的|没问题|非常荣幸|收到|为你生成|以下是|这是为您|正文开始|下面是).*?[:：]\n*', '', text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r'^\s*第[零一二三四五六七八九十百千0-9]+[章回节卷].*?\n', '', text, flags=re.MULTILINE)
    text = re.sub(r'```[a-zA-Z]*\n?', '', text)
    text = re.sub(r'\n*(希望这|如果有需要|请告诉我|期待您的反馈).*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def count_words(text):
    return len(re.findall(r'[\u4e00-\u9fff]', text)) + len(re.findall(r'[a-zA-Z]+', text))

def extract_file_text(uploaded_file):
    name = uploaded_file.name.lower()
    try:
        if name.endswith('.pdf'):
            reader = PdfReader(uploaded_file)
            return "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
        elif name.endswith('.docx'):
            doc = Document(uploaded_file)
            return "\n".join([p.text for p in doc.paragraphs])
        else:
            # 支持 .txt .md .py 等所有文本类文件
            return uploaded_file.getvalue().decode('utf-8', errors='ignore')
    except Exception as e:
        return f"文件解析失败: {str(e)}"

def generate_word_doc(messages, is_pure=False):
    doc = Document()
    doc.add_heading('ZenMux 导出文档', 0)
    for msg in messages:
        if msg["role"] == "system": continue
        if is_pure:
            if msg["role"] == "assistant":
                doc.add_paragraph(clean_novel_text(msg["content"]))
        else:
            doc.add_heading("📌 我" if msg["role"]=="user" else "🤖 AI", level=2)
            doc.add_paragraph(msg["content"])
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def generate_share_html(chat_title, messages):
    """生成精美的HTML分享页面"""
    def md_to_html(text):
        # 转义HTML特殊字符
        text = html.escape(text)
        # 代码块
        text = re.sub(r'```(\w*)\n(.*?)```', r'<pre><code class="lang-\1">\2</code></pre>', text, flags=re.DOTALL)
        # 行内代码
        text = re.sub(r'`([^`]+)`', r'<code>\1</code>', text)
        # 粗体
        text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
        # 斜体
        text = re.sub(r'\*([^*]+)\*', r'<em>\1</em>', text)
        # 标题
        text = re.sub(r'^### (.+)$', r'<h3>\1</h3>', text, flags=re.MULTILINE)
        text = re.sub(r'^## (.+)$', r'<h2>\1</h2>', text, flags=re.MULTILINE)
        text = re.sub(r'^# (.+)$', r'<h1>\1</h1>', text, flags=re.MULTILINE)
        # 换行
        text = text.replace('\n\n', '</p><p>')
        text = text.replace('\n', '<br/>')
        return f'<p>{text}</p>'

    msg_blocks = []
    for msg in messages:
        if msg["role"] == "system": continue
        role_cn = "👤 我" if msg["role"] == "user" else "🤖 AI"
        role_class = "user-msg" if msg["role"] == "user" else "ai-msg"
        avatar_bg = "#667eea" if msg["role"] == "user" else "#10b981"
        content_html = md_to_html(msg["content"])
        word_count = count_words(msg["content"])

        files_html = ""
        if msg.get("files"):
            for f in msg["files"]:
                files_html += f'<div class="file-tag">📎 {html.escape(f["filename"])}</div>'

        msg_blocks.append(f'''
        <div class="message {role_class}">
            <div class="avatar" style="background:{avatar_bg}">{role_cn[0]}</div>
            <div class="bubble">
                <div class="role-name">{role_cn}</div>
                {files_html}
                <div class="content">{content_html}</div>
                <div class="meta">📊 {word_count} 字</div>
            </div>
        </div>
        ''')

    messages_html = "\n".join(msg_blocks)
    export_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    safe_title = html.escape(chat_title)

    template = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{safe_title} - ZenMux 对话分享</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
    font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    padding: 20px;
    color: #2d3748;
}}
.container {{
    max-width: 900px;
    margin: 0 auto;
    background: #fff;
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
}}
.header {{
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff;
    padding: 40px 30px;
    text-align: center;
    position: relative;
}}
.header::before {{
    content: "🐙";
    font-size: 60px;
    display: block;
    margin-bottom: 15px;
    filter: drop-shadow(0 4px 8px rgba(0,0,0,0.2));
}}
.header h1 {{
    font-size: 28px;
    font-weight: 700;
    margin-bottom: 10px;
    text-shadow: 0 2px 4px rgba(0,0,0,0.2);
}}
.header .subtitle {{
    font-size: 14px;
    opacity: 0.9;
}}
.stats-bar {{
    background: #f7fafc;
    padding: 15px 30px;
    border-bottom: 1px solid #e2e8f0;
    display: flex;
    justify-content: space-around;
    flex-wrap: wrap;
    gap: 15px;
}}
.stat-item {{
    text-align: center;
    font-size: 13px;
    color: #718096;
}}
.stat-item .num {{
    display: block;
    font-size: 20px;
    font-weight: 700;
    color: #667eea;
    margin-bottom: 3px;
}}
.messages {{
    padding: 30px;
}}
.message {{
    display: flex;
    margin-bottom: 25px;
    animation: fadeInUp 0.5s ease;
}}
.message.user-msg {{
    flex-direction: row-reverse;
}}
.avatar {{
    width: 44px;
    height: 44px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #fff;
    font-weight: bold;
    font-size: 18px;
    flex-shrink: 0;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}}
.bubble {{
    max-width: 75%;
    margin: 0 15px;
    padding: 18px 22px;
    border-radius: 18px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    position: relative;
}}
.user-msg .bubble {{
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff;
    border-top-right-radius: 4px;
}}
.ai-msg .bubble {{
    background: #f7fafc;
    border: 1px solid #e2e8f0;
    border-top-left-radius: 4px;
}}
.role-name {{
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 8px;
    opacity: 0.85;
}}
.content {{
    line-height: 1.8;
    font-size: 15px;
    word-wrap: break-word;
}}
.content p {{
    margin-bottom: 10px;
}}
.content h1, .content h2, .content h3 {{
    margin: 15px 0 10px;
    font-weight: 700;
}}
.content h1 {{ font-size: 22px; }}
.content h2 {{ font-size: 19px; }}
.content h3 {{ font-size: 17px; }}
.content code {{
    background: rgba(0,0,0,0.1);
    padding: 2px 6px;
    border-radius: 4px;
    font-family: "Monaco", "Menlo", monospace;
    font-size: 13px;
}}
.user-msg .content code {{
    background: rgba(255,255,255,0.25);
}}
.content pre {{
    background: #1a202c;
    color: #e2e8f0;
    padding: 15px;
    border-radius: 8px;
    overflow-x: auto;
    margin: 10px 0;
}}
.content pre code {{
    background: transparent;
    padding: 0;
    color: inherit;
}}
.meta {{
    font-size: 11px;
    opacity: 0.6;
    margin-top: 10px;
    text-align: right;
}}
.file-tag {{
    display: inline-block;
    background: rgba(255,255,255,0.2);
    padding: 4px 10px;
    border-radius: 12px;
    font-size: 12px;
    margin-bottom: 8px;
}}
.ai-msg .file-tag {{
    background: #e6fffa;
    color: #234e52;
}}
.footer {{
    background: #2d3748;
    color: #a0aec0;
    padding: 20px;
    text-align: center;
    font-size: 13px;
}}
.footer a {{
    color: #90cdf4;
    text-decoration: none;
}}
@keyframes fadeInUp {{
    from {{ opacity: 0; transform: translateY(20px); }}
    to {{ opacity: 1; transform: translateY(0); }}
}}
@media (max-width: 600px) {{
    body {{ padding: 10px; }}
    .header {{ padding: 25px 15px; }}
    .header h1 {{ font-size: 22px; }}
    .messages {{ padding: 20px 15px; }}
    .bubble {{ max-width: 85%; padding: 14px 16px; }}
    .avatar {{ width: 36px; height: 36px; font-size: 15px; }}
}}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>{safe_title}</h1>
        <div class="subtitle">ZenMux 创作者工作站 · 对话分享</div>
    </div>
    <div class="stats-bar">
        <div class="stat-item"><span class="num">{len([m for m in messages if m['role']!='system'])}</span>消息数</div>
        <div class="stat-item"><span class="num">{sum(count_words(m['content']) for m in messages if m['role']!='system')}</span>总字数</div>
        <div class="stat-item"><span class="num">{export_time.split(' ')[0]}</span>导出日期</div>
    </div>
    <div class="messages">
        {messages_html}
    </div>
    <div class="footer">
        ✨ 由 <strong>ZenMux 创作者工作站</strong> 生成 · {export_time}
    </div>
</div>
</body>
</html>'''
    return template

def fetch_models(base_url, api_key):
    try:
        url = (base_url.strip().rstrip('/') or "https://api.openai.com/v1") + "/models"
        resp = requests.get(url, headers={"Authorization": "Bearer " + api_key.strip()}, timeout=8)
        if resp.status_code == 200:
            return True, sorted([m["id"] for m in resp.json().get("data", [])])
        return False, f"状态码 {resp.status_code}: {resp.text[:100]}"
    except Exception as e:
        return False, str(e)

def get_client():
    p = st.session_state.profiles[st.session_state.active_profile_idx]
    url = p["base_url"].strip() or "https://api.openai.com/v1"
    return OpenAI(base_url=url, api_key=p["api_key"].strip()), p

def build_api_kwargs(profile, api_msgs):
    kw = {"model": profile["model"], "messages": api_msgs, "stream": True}
    if profile.get("use_temperature", True): kw["temperature"] = profile.get("temperature", 0.8)
    if profile.get("use_max_tokens", True): kw["max_tokens"] = profile.get("max_tokens", 8192)
    if profile.get("use_top_p", False): kw["top_p"] = profile.get("top_p", 1.0)
    if profile.get("use_frequency_penalty", False): kw["frequency_penalty"] = profile.get("frequency_penalty", 0.0)
    return kw

# ==========================================
# 4. 全局侧边栏导航
# ==========================================
with st.sidebar:
    col_img, col_txt = st.columns([1, 3])
    with col_img: st.image("https://api.iconify.design/fluent-emoji:octopus.svg?width=80", width=45)
    with col_txt: st.header("控制中枢")

    pages = ["💬 自由聊天区", "⚙️ 底层引擎配置"]
    for pg in pages:
        btype = "primary" if st.session_state.current_page == pg else "secondary"
        if st.button(pg, use_container_width=True, type=btype):
            st.session_state.current_page = pg
            st.rerun()

    active_p = st.session_state.profiles[st.session_state.active_profile_idx]
    st.divider()
    st.caption(f"🟢 **当前挂载**: {active_p['name']}\n🧠 **模型**: {active_p['model']}")

    st.divider()
    with st.expander("📦 全量资产导出恢复舱", expanded=False):
        st.caption("换电脑时一键导入导出所有数据。")
        full_data = json.dumps({
            "profiles": st.session_state.profiles,
            "free_chats": st.session_state.free_chats
        }, ensure_ascii=False, indent=2).encode('utf-8')
        fname = f"ZenMux_Backup_{datetime.now().strftime('%m%d_%H%M')}.json"
        st.download_button("📥 导出全量快照包", full_data, fname, "application/json", use_container_width=True, type="primary")

        uploaded_ws = st.file_uploader("📂 导入快照 (覆盖当前)", type="json")
        if uploaded_ws:
            try:
                data = json.loads(uploaded_ws.getvalue().decode('utf-8'))
                st.session_state.profiles = data.get("profiles", st.session_state.profiles)
                st.session_state.free_chats = data.get("free_chats", st.session_state.free_chats)
                st.session_state.active_profile_idx = 0
                st.session_state.current_chat_id = list(st.session_state.free_chats.keys())[-1]
                trigger_save()
                st.success("✅ 恢复成功！")
                st.rerun()
            except Exception as e:
                st.error(f"导入失败: {e}")
# ==========================================
# 模块 1: 自由聊天区 (核心主战场)
# ==========================================
if st.session_state.current_page == "💬 自由聊天区":
    # --- 顶部会话管理 ---
    with st.expander("📚 会话列表与管理", expanded=False):
        col_new, col_search = st.columns([1, 2])
        with col_new:
            if st.button("➕ 新对话", use_container_width=True, type="primary"):
                nid = str(uuid.uuid4())
                st.session_state.free_chats[nid] = {"title": "新对话", "messages": [], "session_knowledge": [], "system_prompt": "", "is_pinned": False, "is_archived": False}
                st.session_state.current_chat_id = nid
                trigger_save()
                st.rerun()
        with col_search:
            search_q = st.text_input("🔍 搜索历史对话", label_visibility="collapsed", placeholder="输入关键词搜索...")

        chat_items = []
        for cid, cdata in st.session_state.free_chats.items():
            if cdata.get("is_archived", False): continue
            if search_q and search_q.lower() not in cdata["title"].lower(): continue
            chat_items.append((cid, cdata))
        chat_items.sort(key=lambda x: x[1].get("is_pinned", False), reverse=True)

        for cid, cdata in chat_items:
            prefix = "📌 " if cdata.get("is_pinned") else "📄 "
            if cid == st.session_state.current_chat_id: prefix = "⭐ "
            if st.button(prefix + cdata["title"], key=f"sel_{cid}", use_container_width=True):
                st.session_state.current_chat_id = cid
                st.rerun()

        if st.button("🗄️ 查看归档区", use_container_width=True):
            st.session_state._show_archive = not st.session_state.get("_show_archive", False)
            st.rerun()

        if st.session_state.get("_show_archive", False):
            st.markdown("---")
            st.caption("🗄️ 归档会话")
            for cid, cdata in st.session_state.free_chats.items():
                if cdata.get("is_archived", False):
                    cc1, cc2 = st.columns([4, 1])
                    cc1.button("📦 " + cdata["title"], key=f"arch_{cid}", disabled=True, use_container_width=True)
                    if cc2.button("恢复", key=f"unarch_{cid}"):
                        cdata["is_archived"] = False
                        trigger_save()
                        st.rerun()

    # --- 当前会话主体 ---
    if st.session_state.current_chat_id not in st.session_state.free_chats:
        st.session_state.current_chat_id = list(st.session_state.free_chats.keys())[-1]
    curr_chat = st.session_state.free_chats[st.session_state.current_chat_id]

    # 标题与操作栏
    tc1, tc2, tc3, tc4, tc5, tc6 = st.columns([4, 1, 1, 1, 1, 1])
    with tc1:
        new_title = st.text_input("会话标题", curr_chat["title"], label_visibility="collapsed")
        if new_title != curr_chat["title"]:
            curr_chat["title"] = new_title
            trigger_save()
    with tc2:
        pin_lbl = "取消置顶" if curr_chat.get("is_pinned") else "📌 置顶"
        if st.button(pin_lbl, use_container_width=True):
            curr_chat["is_pinned"] = not curr_chat.get("is_pinned", False)
            trigger_save()
            st.rerun()
    with tc3:
        if st.button("📦 归档", use_container_width=True):
            curr_chat["is_archived"] = True
            trigger_save()
            st.rerun()
    with tc4:
        if st.button("🗑️ 清空", use_container_width=True):
            curr_chat["messages"] = []
            trigger_save()
            st.rerun()
    with tc5:
        if st.button("📥 导出", use_container_width=True):
            st.session_state._show_export = not st.session_state.get("_show_export", False)
            st.rerun()
    with tc6:
        if st.button("🔗 分享", use_container_width=True):
            st.session_state._show_share = not st.session_state.get("_show_share", False)
            st.rerun()

    # 导出面板
    if st.session_state.get("_show_export", False):
        with st.container(border=True):
            st.markdown("**📦 导出当前对话**")
            exp_mode = st.radio("格式", ["完整记录", "纯享正文"], horizontal=True, label_visibility="collapsed")
            is_pure = (exp_mode == "纯享正文")
            curr_msgs = curr_chat["messages"]

            if is_pure:
                txt_c = "\n\n".join([clean_novel_text(m['content']) for m in curr_msgs if m['role'] == 'assistant'])
            else:
                txt_c = "\n".join([f"{'我' if m['role']=='user' else 'AI'}:\n{m['content']}\n\n{'-'*40}\n" for m in curr_msgs])

            ec1, ec2 = st.columns(2)
            ec1.download_button("📥 下载 TXT", txt_c.encode('utf-8'), f"{curr_chat['title']}.txt", use_container_width=True)
            ec2.download_button("📥 下载 Word", generate_word_doc(curr_msgs, is_pure), f"{curr_chat['title']}.docx", use_container_width=True)

    # 分享HTML面板
    if st.session_state.get("_show_share", False):
        with st.container(border=True):
            st.markdown("**🔗 生成精美HTML分享页**")
            st.caption("生成一份独立的、排版精美的HTML文件，可直接在任意浏览器打开，或上传到网盘/网站分享。")
            share_html = generate_share_html(curr_chat["title"], curr_chat["messages"])
            sc1, sc2 = st.columns(2)
            sc1.download_button(
                "🎨 下载精美HTML分享页", 
                share_html.encode('utf-8'), 
                f"分享_{curr_chat['title']}_{datetime.now().strftime('%m%d')}.html", 
                "text/html",
                use_container_width=True,
                type="primary"
            )
            if sc2.button("👀 预览分享效果", use_container_width=True):
                st.session_state._show_preview = not st.session_state.get("_show_preview", False)
                st.rerun()

            if st.session_state.get("_show_preview", False):
                components.html(share_html, height=600, scrolling=True)

    # --- 对话设置区 ---
    has_content = bool(curr_chat.get("session_knowledge") or curr_chat.get("system_prompt"))
    with st.expander("⚙️ 全局设定与常驻知识库", expanded=has_content):
        curr_chat["system_prompt"] = st.text_area("🎭 System Prompt (全局人设)", curr_chat.get("system_prompt", ""), height=80)
        up_f = st.file_uploader(
            "📎 上传常驻参考文件 (TXT/MD/PDF/DOCX/PY/JSON/HTML/CSS/JS)", 
            type=['txt', 'md', 'pdf', 'docx', 'py', 'json', 'html', 'css', 'js', 'xml', 'yaml', 'yml', 'csv', 'log', 'ini', 'cfg', 'sh', 'bat', 'java', 'cpp', 'c', 'h', 'go', 'rs', 'ts', 'jsx', 'tsx', 'vue', 'php', 'rb', 'sql'], 
            key=f"kb_{st.session_state.current_chat_id}"
        )
        if up_f:
            content = extract_file_text(up_f)
            if not any(k["filename"] == up_f.name for k in curr_chat.get("session_knowledge", [])):
                if "session_knowledge" not in curr_chat: curr_chat["session_knowledge"] = []
                curr_chat["session_knowledge"].append({"filename": up_f.name, "content": content})
                trigger_save()
                st.rerun()
        for ki, k in enumerate(curr_chat.get("session_knowledge", [])):
            kc1, kc2 = st.columns([4, 1])
            kc1.caption(f"📄 {k['filename']} ({count_words(k['content']):,} 字)")
            if kc2.button("❌", key=f"rm_kb_{ki}"):
                curr_chat["session_knowledge"].pop(ki)
                trigger_save()
                st.rerun()

    # --- 聊天消息展示 ---
    with st.container(height=550, border=False):
        editing_idx = st.session_state.get("_editing_chat_idx")

        for i, msg in enumerate(curr_chat["messages"]):
            if msg["role"] == "system": continue

            with st.chat_message(msg["role"]):
                if msg.get("files"):
                    for f in msg["files"]:
                        badge = "🔄" if f.get("continuous") else "1️⃣"
                        st.caption(f"`{badge} 附件: {f['filename']}`")

                if msg["role"] == "user" and editing_idx == i:
                    new_text = st.text_area("✏️ 编辑消息", msg["content"], key=f"edit_area_{i}", height=100)
                    ebc1, ebc2 = st.columns(2)
                    if ebc1.button("✅ 确认并重新发送", key=f"edit_ok_{i}", type="primary"):
                        msg["content"] = new_text
                        curr_chat["messages"] = curr_chat["messages"][:i + 1]
                        st.session_state._editing_chat_idx = None
                        st.session_state._auto_resend = True
                        trigger_save()
                        st.rerun()
                    if ebc2.button("❌ 取消", key=f"edit_cancel_{i}"):
                        st.session_state._editing_chat_idx = None
                        st.rerun()
                else:
                    st.markdown(msg["content"])

                    mc1, mc2, mc3, mc4 = st.columns([1, 1, 1, 5])
                    if msg["role"] == "user" and editing_idx is None:
                        if mc1.button("✏️", key=f"edit_btn_{i}", help="编辑"):
                            st.session_state._editing_chat_idx = i
                            st.rerun()
                    if msg["role"] == "assistant":
                        if mc1.button("🔄", key=f"regen_{i}", help="重新生成"):
                            curr_chat["messages"] = curr_chat["messages"][:i]
                            st.session_state._auto_resend = True
                            trigger_save()
                            st.rerun()
                        if msg.get("_is_half"):
                            if mc2.button("▶️ 续写", key=f"resume_{i}", help="从中断处继续"):
                                st.session_state._resume_idx = i
                                st.session_state._auto_resend = True
                                st.rerun()
                    if mc3.button("🗑️", key=f"del_msg_{i}", help="删除此条"):
                        curr_chat["messages"].pop(i)
                        trigger_save()
                        st.rerun()

                    if msg["role"] == "assistant":
                        render_copy_button(msg["content"])
                        st.caption(f"📊 {count_words(msg['content'])} 字")

    # --- 动态文件挂载与输入区 ---
    need_resend = st.session_state.pop("_auto_resend", False)
    resume_idx = st.session_state.pop("_resume_idx", None)

    with st.expander("📎 随消息挂载单次/持续附件", expanded=False):
        dyn_file = st.file_uploader(
            "上传文件 (TXT/MD/PDF/DOCX/PY/代码文件等)", 
            type=['txt', 'md', 'pdf', 'docx', 'py', 'json', 'html', 'css', 'js', 'xml', 'yaml', 'yml', 'csv', 'log', 'ini', 'cfg', 'sh', 'bat', 'java', 'cpp', 'c', 'h', 'go', 'rs', 'ts', 'jsx', 'tsx', 'vue', 'php', 'rb', 'sql'], 
            key="dyn_file"
        )
        is_continuous = st.checkbox("🔄 持续参考 (开启后该文件将一直带入后续对话，否则仅本次有效)", value=False)

    prompt = st.chat_input("输入消息...")

    # ==========================================
    # 🔥 核心修复：流式请求逻辑（修复消息消失问题）
    # ==========================================
    if prompt or need_resend:
        if not active_p["api_key"]:
            st.error("⚠️ 请先在左侧【底层引擎配置】中填写 API Key！")
            st.stop()

        if resume_idx is not None:
            prompt = "请紧接上文最后一个字继续往下写，保持文风和节奏一致。"
            curr_chat["messages"][resume_idx]["_is_half"] = False

        if prompt:
            if len(curr_chat["messages"]) == 0 and curr_chat["title"] == "新对话":
                curr_chat["title"] = prompt[:10] + ("..." if len(prompt)>10 else "")

            new_msg = {"role": "user", "content": prompt}
            if dyn_file and not need_resend:
                new_msg["files"] = [{
                    "filename": dyn_file.name,
                    "content": extract_file_text(dyn_file),
                    "continuous": is_continuous
                }]
            curr_chat["messages"].append(new_msg)

        if prompt:
            with st.chat_message("user"):
                st.markdown(prompt)

        # 构建 API 消息包
        api_msgs = []
        sp = curr_chat.get("system_prompt", "").strip()
        if sp: api_msgs.append({"role": "system", "content": sp})

        if curr_chat.get("session_knowledge"):
            kb_parts = [f"--- 文件: {k['filename']} ---\n{k['content']}" for k in curr_chat["session_knowledge"]]
            api_msgs.append({"role": "system", "content": "【全局参考文件】：\n" + "\n\n".join(kb_parts)})

        for i, m in enumerate(curr_chat["messages"]):
            content = m["content"]
            if m.get("files"):
                file_texts = []
                for f in m["files"]:
                    if not f.get("continuous") and i != len(curr_chat["messages"]) - 1:
                        continue
                    file_texts.append(f"--- 附件: {f['filename']} ---\n{f['content']}")
                if file_texts:
                    content = "【参考附件】\n" + "\n".join(file_texts) + "\n\n【用户指令】\n" + content

            api_msgs.append({"role": m["role"], "content": content})

        # 🔥 关键修复：正确的流式输出 + 防消失
        client, profile = get_client()

        # 先标记流式状态，阻止execute_save干扰
        st.session_state.is_streaming = True

        with st.chat_message("assistant"):
            message_placeholder = st.empty()
            full_resp = ""
            is_half = False
            error_occurred = False
            error_msg = ""

            try:
                resp = client.chat.completions.create(**build_api_kwargs(profile, api_msgs))
                for chunk in resp:
                    if chunk.choices and chunk.choices[0].delta.content is not None:
                        full_resp += chunk.choices[0].delta.content
                        # 降低刷新频率，每32字符刷新一次，避免页面卡顿
                        if len(full_resp) % 32 == 0:
                            message_placeholder.markdown(full_resp + "▌")

                message_placeholder.markdown(full_resp)

            except Exception as e:
                error_occurred = True
                error_msg = str(e)
                is_half = True
                if full_resp:
                    message_placeholder.markdown(full_resp)

            # 🔥 无论成功失败，都必须保存！防止消息丢失
            if full_resp:
                curr_chat["messages"].append({
                    "role": "assistant", 
                    "content": full_resp, 
                    "_is_half": is_half
                })

            # 立即写入LocalStorage（先解除流式锁）
            st.session_state.is_streaming = False
            st.session_state._needs_save = True
            execute_save()

            if error_occurred:
                st.error(f"⚠️ 请求中断: {error_msg}")
                if full_resp:
                    st.info("✅ 已生成的内容已保留，可点击消息下方的【▶️ 续写】按钮继续。")

            # 不再立即rerun，等用户下一次交互再刷新，避免打断
            if not error_occurred:
                st.rerun()

# ==========================================
# 模块 2: 底层引擎配置
# ==========================================
elif st.session_state.current_page == "⚙️ 底层引擎配置":
    st.header("⚙️ 底层驱动配置")
    col_list, col_edit = st.columns([1, 2.5])

    with col_list:
        st.subheader("引擎库")
        p_names = [p["name"] for p in st.session_state.profiles]
        idx = st.radio("切换引擎", range(len(p_names)), format_func=lambda x: p_names[x], index=st.session_state.active_profile_idx)
        st.session_state.active_profile_idx = idx

        if st.button("➕ 新增引擎", use_container_width=True):
            st.session_state.profiles.append({
                "name": f"新引擎 {len(p_names) + 1}", "base_url": "", "api_key": "", "model": "anthropic/claude-3-5-sonnet-20240620",
                "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 8192,
                "use_top_p": False, "top_p": 1.0, "use_frequency_penalty": False, "frequency_penalty": 0.0
            })
            trigger_save()
            st.rerun()

    with col_edit:
        st.subheader("网络与参数调优")
        p = st.session_state.profiles[idx]

        eca, ecb = st.columns([3, 1])
        with eca: p["name"] = st.text_input("引擎标签", p["name"])
        with ecb:
            st.write("")
            if st.button("💾 保存配置", type="primary", use_container_width=True):
                trigger_save()
                st.success("已保存！")

        nc1, nc2 = st.columns(2)
        with nc1: p["base_url"] = st.text_input("Base URL", p["base_url"])
        with nc2: 
            p["api_key"] = st.text_input("API Key", p["api_key"], type="password")
            st.caption("🔒 **安全提示**：你的 Key 只保存在你自己的浏览器本地缓存中，服务器不会记录。")

        if p["api_key"]:
            if st.button("🔑 测试连通性"):
                with st.spinner("正在测试..."):
                    try:
                        test_client = OpenAI(base_url=p["base_url"].strip() or "https://api.openai.com/v1", api_key=p["api_key"].strip())
                        test_client.chat.completions.create(model=p["model"], messages=[{"role": "user", "content": "Hi"}], max_tokens=5)
                        st.success("✅ 连通成功！模型响应正常。")
                    except Exception as e:
                        st.error(f"❌ 连通失败: {str(e)}")

        cm, cb2 = st.columns([3, 1])
        with cm: p["model"] = st.text_input("模型映射 (Model ID)", p["model"])
        with cb2:
            st.write("")
            if st.button("🔄 联机获取列表"):
                if p["api_key"]:
                    with st.spinner("正在获取..."):
                        success, result = fetch_models(p["base_url"], p["api_key"])
                        if success and result:
                            st.session_state.temp_models = result
                            st.success(f"✅ 抓取到 {len(result)} 个模型！")
                        else:
                            st.error(f"❌ 获取失败: {result}")
                else:
                    st.error("请先填写 API Key！")

        if "temp_models" in st.session_state:
            sel_m = st.selectbox("选择模型", ["(不覆盖)"] + st.session_state.temp_models)
            if sel_m != "(不覆盖)":
                p["model"] = sel_m
                del st.session_state.temp_models
                trigger_save()
                st.rerun()

        st.markdown("#### 🎛️ 运行时超参数 (勾选生效)")
        sl1, sl2 = st.columns(2)
        with sl1:
            p["use_temperature"] = st.checkbox("🔥 Temperature", p.get("use_temperature", True))
            if p["use_temperature"]: 
                p["temperature"] = st.slider("温度值", 0.0, 2.0, p.get("temperature", 0.8), 0.1, label_visibility="collapsed")

            p["use_max_tokens"] = st.checkbox("📏 Max Tokens (最大输出长度)", p.get("use_max_tokens", True))
            if p["use_max_tokens"]: 
                # 🔥 关键改进：提供输入框+滑块双模式，支持最高200K
                mt_mode = st.radio(
                    "Max Tokens 设置方式", 
                    ["快捷档位", "精确输入"], 
                    horizontal=True, 
                    key=f"mt_mode_{idx}",
                    label_visibility="collapsed"
                )
                if mt_mode == "快捷档位":
                    preset_options = {
                        "4K (默认)": 4096,
                        "8K": 8192,
                        "16K": 16384,
                        "32K (Claude)": 32768,
                        "64K (GPT-4o)": 65536,
                        "128K (Gemini)": 131072,
                        "200K (Claude 3.5 Max)": 200000,
                    }
                    current_val = p.get("max_tokens", 8192)
                    # 找到最接近的档位
                    default_key = min(preset_options.keys(), key=lambda k: abs(preset_options[k] - current_val))
                    selected = st.selectbox(
                        "选择档位", 
                        list(preset_options.keys()), 
                        index=list(preset_options.keys()).index(default_key),
                        key=f"mt_preset_{idx}",
                        label_visibility="collapsed"
                    )
                    p["max_tokens"] = preset_options[selected]
                    st.caption(f"当前值: **{p['max_tokens']:,}** tokens")
                else:
                    new_mt = st.number_input(
                        "精确输入 (建议范围: 512 ~ 200000)", 
                        min_value=128, 
                        max_value=500000,  # 上限放到50万，覆盖所有模型
                        value=p.get("max_tokens", 8192), 
                        step=512,
                        key=f"mt_num_{idx}",
                        help="⚠️ 不同模型上限不同：GPT-4o=16K输出，Claude 3.5=8K输出，Gemini 2.0=8K输出。超过会报错，请按模型文档设置。"
                    )
                    p["max_tokens"] = new_mt
                    st.caption(f"💡 当前值: **{p['max_tokens']:,}** tokens | 注意：这是最大**输出**长度，不是上下文窗口长度")

        with sl2:
            p["use_top_p"] = st.checkbox("🎲 Top P", p.get("use_top_p", False))
            if p["use_top_p"]: 
                p["top_p"] = st.slider("Top P值", 0.0, 1.0, p.get("top_p", 1.0), 0.05, label_visibility="collapsed")
            p["use_frequency_penalty"] = st.checkbox("🚫 Frequency Penalty", p.get("use_frequency_penalty", False))
            if p["use_frequency_penalty"]: 
                p["frequency_penalty"] = st.slider("惩罚值", -2.0, 2.0, p.get("frequency_penalty", 0.0), 0.1, label_visibility="collapsed")

        if len(st.session_state.profiles) > 1:
            st.divider()
            if st.button("🗑️ 删除此引擎"):
                st.session_state.profiles.pop(idx)
                st.session_state.active_profile_idx = 0
                trigger_save()
                st.rerun()

# 统一执行保存
execute_save()
