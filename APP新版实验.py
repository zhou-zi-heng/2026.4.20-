import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
import io, base64, json, re, requests, uuid, copy
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
    /* 内联复制按钮样式 */
    .zm-copy-btn {
        border:none;background:transparent;color:#aaa;cursor:pointer;font-size:12px;
        font-weight:bold;padding:4px 10px;border-radius:6px;float:right;
    }
    .zm-copy-btn:hover { background:rgba(255,255,255,0.08); color:#fff; }
    .zm-word-tag { color:#888; font-size:12px; margin-left:8px; }
    </style>
    <script>
    // 全局单例复制函数（只注入一次，替代 N 个 iframe）
    window.zmCopy = function(btn, b64) {
        try {
            const text = decodeURIComponent(escape(atob(b64)));
            navigator.clipboard.writeText(text).then(function(){
                const old = btn.innerText;
                btn.innerText = '✅ 已复制';
                btn.style.color = '#4CAF50';
                setTimeout(function(){ btn.innerText = old; btn.style.color = '#aaa'; }, 2000);
            });
        } catch(e) { alert('复制失败：' + e); }
    };
    </script>
""", unsafe_allow_html=True)

# ==========================================
# 2. 数据持久化层 (LocalStorage 桥接)
# ==========================================
localS = LocalStorage()

# 支持的文件类型（集中管理，便于后续扩展）
SUPPORTED_FILE_TYPES = ['txt', 'md', 'py', 'json', 'csv', 'log', 'pdf', 'docx']
FILE_TYPE_HINT = "TXT/MD/PY/JSON/CSV/LOG/PDF/DOCX"

# 长会话渲染分页
DEFAULT_RENDER_WINDOW = 50  # 默认只渲染最后 N 条消息

def trigger_save():
    """标记需要保存，避免在流式输出时频繁触发 rerun"""
    st.session_state._needs_save = True

def execute_save():
    """真正执行写入 LocalStorage"""
    if st.session_state.get("_needs_save", False) and not st.session_state.get("is_streaming", False):
        data = {
            "profiles": st.session_state.profiles,
            "free_chats": st.session_state.free_chats
        }
        localS.setItem("zenmux_data", json.dumps(data))
        st.session_state._needs_save = False

# 初始化状态
if "initialized" not in st.session_state:
    st.session_state.initialized = False
    st.session_state.ls_loaded = False
    st.session_state._needs_save = False
    st.session_state.is_streaming = False
    st.session_state.stop_req = False
    st.session_state._stream_content = ""
    st.session_state._stream_chat_id = None
    st.session_state._render_limit = DEFAULT_RENDER_WINDOW

# 水合逻辑 (从 LocalStorage 读取) —— 修复版
if not st.session_state.ls_loaded:
    # 默认数据（无论如何都要准备好）
    default_profiles = [{
        "name": "默认引擎", "base_url": "", "api_key": "", "model": "anthropic/claude-3-5-sonnet-20240620",
        "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 4096,
        "use_top_p": False, "top_p": 1.0, "use_frequency_penalty": False, "frequency_penalty": 0.0
    }]
    first_id = str(uuid.uuid4())
    default_chats = {first_id: {"title": "新对话", "messages": [], "session_knowledge": [],
                                 "system_prompt": "", "is_pinned": False, "is_archived": False}}

    saved_data = localS.getItem("zenmux_data")

    # 计数重试：前端组件首次挂载需要 1~2 次 rerun 才能回传真实值
    retry = st.session_state.get("_ls_retry", 0)

    if saved_data is None and retry < 3:
        # 还没拿到，继续等；用占位符提示并自动重试
        st.session_state._ls_retry = retry + 1
        placeholder = st.empty()
        placeholder.info(f"🔄 正在从本地安全存储加载数据... ({retry + 1}/3)")
        import time
        time.sleep(0.4)
        placeholder.empty()
        st.rerun()

    # 到这里：要么拿到 saved_data，要么重试完仍为 None（视为全新用户）
    try:
        if saved_data and isinstance(saved_data, str):
            data = json.loads(saved_data)
            st.session_state.profiles = data.get("profiles", default_profiles)
            st.session_state.free_chats = data.get("free_chats", default_chats)
        elif saved_data and isinstance(saved_data, dict):
            st.session_state.profiles = saved_data.get("profiles", default_profiles)
            st.session_state.free_chats = saved_data.get("free_chats", default_chats)
        else:
            # 全新用户或读取失败，使用默认值
            st.session_state.profiles = default_profiles
            st.session_state.free_chats = default_chats
    except Exception as e:
        st.warning(f"⚠️ 本地数据解析失败，已使用默认配置：{e}")
        st.session_state.profiles = default_profiles
        st.session_state.free_chats = default_chats

    # 容错：确保至少有一个对话
    if not st.session_state.free_chats:
        st.session_state.free_chats = default_chats

    st.session_state.active_profile_idx = 0
    st.session_state.current_chat_id = list(st.session_state.free_chats.keys())[-1]
    st.session_state.current_page = "💬 自由聊天区"
    st.session_state.ls_loaded = True
    st.session_state.initialized = True
    st.session_state._ls_retry = 0
    st.rerun()

# ==========================================
# 3. 核心底层辅助函数
# ==========================================
def render_copy_button_inline(text):
    """内联复制按钮（不使用 iframe，大幅减少渲染开销）"""
    b64 = base64.b64encode(text.encode("utf-8")).decode("utf-8")
    html = (
        f'<button class="zm-copy-btn" onclick="window.zmCopy(this, \'{b64}\')">📋 复制</button>'
    )
    return html

def clean_novel_text(text):
    text = re.sub(r'^\s*(好的|没问题|非常荣幸|收到|为你生成|以下是|这是为您|正文开始|下面是).*?[:：]\n*', '', text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r'^\s*第[零一二三四五六七八九十百千0-9]+[章回节卷].*?\n', '', text, flags=re.MULTILINE)
    text = re.sub(r'```[a-zA-Z]*\n?', '', text)
    text = re.sub(r'\n*(希望这|如果有需要|请告诉我|期待您的反馈).*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

@st.cache_data(show_spinner=False, max_entries=2000)
def count_words(text):
    """字数统计（缓存结果，避免 rerun 时重复计算）"""
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
            # 覆盖 txt/md/py/json/csv/log 等所有纯文本格式
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
    if profile.get("use_max_tokens", True):
        kw["max_tokens"] = profile.get("max_tokens", 4096) or 4096
    if profile.get("use_top_p", False): kw["top_p"] = profile.get("top_p", 1.0)
    if profile.get("use_frequency_penalty", False): kw["frequency_penalty"] = profile.get("frequency_penalty", 0.0)
    return kw

def build_share_html(chat):
    """生成精美的 HTML 分享卡片"""
    title = chat.get("title", "对话记录")
    messages = chat.get("messages", [])
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    html = f"""<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title} - ZenMux 对话分享</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background: #1a1a2e; margin: 0; padding: 2rem 1rem; color: #eee; line-height: 1.6;
        }}
        .container {{ max-width: 800px; margin: 0 auto; }}
        .header {{ text-align: center; margin-bottom: 2rem; padding-bottom: 1rem; border-bottom: 2px solid #16213e; }}
        .header h1 {{ background: linear-gradient(135deg, #f6d365 0%, #fda085 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-size: 2rem; }}
        .header .meta {{ color: #a0a0a0; font-size: 0.9rem; }}
        .message {{ display: flex; margin-bottom: 1.5rem; animation: fadeIn 0.3s ease; }}
        @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(10px); }} to {{ opacity: 1; transform: translateY(0); }} }}
        .message.user {{ justify-content: flex-end; }}
        .bubble {{ max-width: 80%; padding: 0.8rem 1.2rem; border-radius: 18px; position: relative; word-break: break-word; white-space: pre-wrap; }}
        .user .bubble {{ background: #2b5876; background: linear-gradient(135deg, #2b5876 0%, #4e4376 100%); color: white; border-bottom-right-radius: 4px; }}
        .assistant .bubble {{ background: #16213e; color: #e0e0e0; border-bottom-left-radius: 4px; box-shadow: 0 2px 10px rgba(0,0,0,0.3); }}
        .role {{ font-size: 0.75rem; font-weight: bold; margin-bottom: 0.3rem; opacity: 0.8; }}
        .system {{ background: #0f3460; color: #e94560; font-style: italic; margin: 1rem 0; padding: 0.8rem; border-radius: 8px; text-align: center; }}
        .footer {{ text-align: center; margin-top: 2rem; padding-top: 1rem; border-top: 1px solid #16213e; color: #666; font-size: 0.8rem; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📜 {title}</h1>
            <div class="meta">由 ZenMux 生成 · {date_str} · 共 {len(messages)} 轮对话</div>
        </div>
"""
    for msg in messages:
        role = msg.get("role", "user")
        content = msg.get("content", "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if role == "system":
            html += f"<div class='system'>⚙️ {content}</div>"
            continue
        role_class = "user" if role == "user" else "assistant"
        display_role = "🧑 我" if role == "user" else "🤖 AI"
        html += f"""
        <div class="message {role_class}">
            <div class="bubble">
                <div class="role">{display_role}</div>
                {content}
            </div>
        </div>
"""
    html += """
        <div class="footer">
            ✨ 由 ZenMux 创作者工作站生成 · 纯本地运行，安全私密
        </div>
    </div>
</body>
</html>
"""
    return html

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

    # 全量资产导出恢复舱（懒计算：只在展开时才 json.dumps）
    st.divider()
    with st.expander("📦 全量资产导出恢复舱", expanded=False):
        st.caption("换电脑时一键导入导出所有数据。")
        # 用户展开后才真正生成导出包（避免每次 rerun 都序列化大 JSON）
        if st.button("🔧 生成快照", use_container_width=True):
            st.session_state._snapshot_ready = json.dumps({
                "profiles": st.session_state.profiles,
                "free_chats": st.session_state.free_chats
            }, ensure_ascii=False, indent=2).encode('utf-8')
        if st.session_state.get("_snapshot_ready"):
            fname = f"ZenMux_Backup_{datetime.now().strftime('%m%d_%H%M')}.json"
            st.download_button("📥 下载快照包", st.session_state._snapshot_ready, fname,
                               "application/json", use_container_width=True, type="primary")

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
# 模块 1: 自由聊天区
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
                st.session_state._render_limit = DEFAULT_RENDER_WINDOW
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
                st.session_state._render_limit = DEFAULT_RENDER_WINDOW
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
    tc1, tc2, tc3, tc4, tc5, tc6 = st.columns([3, 1, 1, 1, 1, 1])
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
        if st.button("📤 分享", use_container_width=True):
            st.session_state._show_share = not st.session_state.get("_show_share", False)
    with tc6:
        if st.button("📥 导出", use_container_width=True):
            st.session_state._show_export = not st.session_state.get("_show_export", False)

    # 分享区域
    if st.session_state.get("_show_share", False):
        with st.container(border=True):
            st.markdown("### 📤 精美 HTML 分享卡")
            share_html = build_share_html(curr_chat)
            components.html(share_html, height=400, scrolling=True)
            col_down, col_copy = st.columns(2)
            with col_down:
                st.download_button("💾 下载 HTML 文件", share_html, f"{curr_chat['title']}_share.html", "text/html", use_container_width=True)
            with col_copy:
                # 使用内联复制按钮（无 iframe）
                st.markdown(
                    f'<div style="text-align:center;padding-top:4px">'
                    + render_copy_button_inline(share_html)
                    + '</div>',
                    unsafe_allow_html=True
                )

    # 导出区域
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

    # --- 对话设置区 (常驻知识库) ---
    has_content = bool(curr_chat.get("session_knowledge") or curr_chat.get("system_prompt"))
    with st.expander("⚙️ 全局设定与常驻知识库", expanded=has_content):
        curr_chat["system_prompt"] = st.text_area("🎭 System Prompt (全局人设)", curr_chat.get("system_prompt", ""), height=80)
        up_f = st.file_uploader(
            f"📎 上传常驻参考文件 ({FILE_TYPE_HINT})",
            type=SUPPORTED_FILE_TYPES,
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

    # --- 聊天消息展示（优化：精简按钮 + 分页渲染 + 内联复制） ---
    all_msgs = curr_chat["messages"]
    total_msgs = len(all_msgs)
    render_limit = st.session_state.get("_render_limit", DEFAULT_RENDER_WINDOW)

    # 计算要渲染的切片（只渲染最后 N 条）
    if total_msgs > render_limit:
        start_idx = total_msgs - render_limit
    else:
        start_idx = 0

    # 找出"最后一条 user / 最后一条 assistant"的全局索引，用于决定按钮显示
    last_user_idx = -1
    last_assistant_idx = -1
    for _i in range(total_msgs - 1, -1, -1):
        _r = all_msgs[_i]["role"]
        if _r == "user" and last_user_idx == -1:
            last_user_idx = _i
        elif _r == "assistant" and last_assistant_idx == -1:
            last_assistant_idx = _i
        if last_user_idx != -1 and last_assistant_idx != -1:
            break

    with st.container(height=550, border=False):
        editing_idx = st.session_state.get("_editing_chat_idx")

        # 顶部"加载更早消息"按钮
        if start_idx > 0:
            lc1, lc2, lc3 = st.columns([1, 2, 1])
            with lc2:
                if st.button(f"📜 加载更早的消息（还剩 {start_idx} 条）", use_container_width=True, key="load_more_msgs"):
                    st.session_state._render_limit = render_limit + DEFAULT_RENDER_WINDOW
                    st.rerun()
            st.caption(f"— 已省略 {start_idx} 条历史消息以保持流畅 —")

        for i in range(start_idx, total_msgs):
            msg = all_msgs[i]
            if msg["role"] == "system": continue

            # 判断是否为"最近一条"——只有最近一条才显示完整操作按钮
            is_last_user = (i == last_user_idx)
            is_last_assistant = (i == last_assistant_idx)

            with st.chat_message(msg["role"]):
                if msg.get("files"):
                    for f in msg["files"]:
                        badge = "🔄" if f.get("continuous") else "1️⃣"
                        st.caption(f"`{badge} 附件: {f['filename']}`")

                # 编辑模式
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

                    # 精简按钮策略：
                    # - 最后一条 user: ✏️编辑 + 🗑️删除
                    # - 最后一条 assistant: 🔄重生 + (可选)▶️续写 + 🗑️删除 + 📋复制 + 字数
                    # - 其他历史消息: 只保留 🗑️删除；assistant 保留 📋复制 + 字数
                    if msg["role"] == "user":
                        if is_last_user and editing_idx is None:
                            mc1, mc2, _sp = st.columns([1, 1, 6])
                            if mc1.button("✏️", key=f"edit_btn_{i}", help="编辑"):
                                st.session_state._editing_chat_idx = i
                                st.rerun()
                            if mc2.button("🗑️", key=f"del_msg_{i}", help="删除此条"):
                                curr_chat["messages"].pop(i)
                                trigger_save()
                                st.rerun()
                        else:
                            # 历史 user 消息只保留删除（小按钮）
                            mc1, _sp = st.columns([1, 9])
                            if mc1.button("🗑️", key=f"del_msg_{i}", help="删除此条"):
                                curr_chat["messages"].pop(i)
                                trigger_save()
                                st.rerun()

                    else:  # assistant
                        if is_last_assistant:
                            # 最后一条 AI：完整按钮
                            has_resume = msg.get("_is_half", False)
                            if has_resume:
                                mc1, mc2, mc3, _sp = st.columns([1, 1, 1, 6])
                            else:
                                mc1, mc3, _sp = st.columns([1, 1, 8])
                                mc2 = None

                            if mc1.button("🔄", key=f"regen_{i}", help="重新生成"):
                                curr_chat["messages"] = curr_chat["messages"][:i]
                                st.session_state._auto_resend = True
                                trigger_save()
                                st.rerun()
                            if mc2 is not None and has_resume:
                                if mc2.button("▶️", key=f"resume_{i}", help="从中断处续写"):
                                    st.session_state._resume_idx = i
                                    st.session_state._auto_resend = True
                                    st.rerun()
                            if mc3.button("🗑️", key=f"del_msg_{i}", help="删除此条"):
                                curr_chat["messages"].pop(i)
                                trigger_save()
                                st.rerun()
                        else:
                            # 历史 AI 消息只保留删除
                            mc1, _sp = st.columns([1, 9])
                            if mc1.button("🗑️", key=f"del_msg_{i}", help="删除此条"):
                                curr_chat["messages"].pop(i)
                                trigger_save()
                                st.rerun()

                        # 复制按钮 + 字数（内联 HTML，无 iframe）
                        wc = count_words(msg["content"])
                        st.markdown(
                            f'<div style="display:flex;justify-content:flex-end;align-items:center;margin-top:4px;">'
                            f'<span class="zm-word-tag">📊 {wc} 字</span>'
                            f'{render_copy_button_inline(msg["content"])}'
                            f'</div>',
                            unsafe_allow_html=True
                        )

    # --- 动态文件挂载与输入区 ---
    need_resend = st.session_state.pop("_auto_resend", False)
    resume_idx = st.session_state.pop("_resume_idx", None)

    with st.expander("📎 随消息挂载单次/持续附件", expanded=False):
        dyn_file = st.file_uploader(
            f"上传文件 ({FILE_TYPE_HINT})",
            type=SUPPORTED_FILE_TYPES,
            key="dyn_file"
        )
        is_continuous = st.checkbox("🔄 持续参考 (开启后该文件将一直带入后续对话，否则仅本次有效)", value=False)

    # --- 发送与流式请求逻辑 ---
    if prompt := st.chat_input("输入消息..."):
        need_resend = False
        if not active_p["api_key"]:
            st.error("⚠️ 请先在左侧【底层引擎配置】中填写 API Key！")
            st.stop()

        # 处理续写逻辑
        if resume_idx is not None:
            prompt = "请紧接上文最后一个字继续往下写，保持文风和节奏一致。"
            curr_chat["messages"][resume_idx]["_is_half"] = False

        if prompt:
            if len(curr_chat["messages"]) == 0 and curr_chat["title"] == "新对话":
                curr_chat["title"] = prompt[:10] + ("..." if len(prompt)>10 else "")

            new_msg = {"role": "user", "content": prompt}
            if dyn_file:
                new_msg["files"] = [{
                    "filename": dyn_file.name,
                    "content": extract_file_text(dyn_file),
                    "continuous": is_continuous
                }]
            curr_chat["messages"].append(new_msg)
            trigger_save()

        if prompt:
            with st.chat_message("user"):
                st.markdown(prompt)

    # 统一的"需要重新生成"处理
    if need_resend or (st.session_state.get("_stream_content") and st.session_state.get("_stream_chat_id") == st.session_state.current_chat_id):
        if not active_p["api_key"]:
            st.error("⚠️ 请先配置 API Key！")
            st.stop()

        curr_msgs = curr_chat["messages"]
        if not curr_msgs:
            st.warning("没有消息可生成。")
        else:
            api_msgs = []
            sp = curr_chat.get("system_prompt", "").strip()
            if sp: api_msgs.append({"role": "system", "content": sp})

            if curr_chat.get("session_knowledge"):
                kb_parts = [f"--- 文件: {k['filename']} ---\n{k['content']}" for k in curr_chat["session_knowledge"]]
                api_msgs.append({"role": "system", "content": "【全局参考文件】：\n" + "\n\n".join(kb_parts)})

            for i, m in enumerate(curr_msgs):
                content = m["content"]
                if m.get("files"):
                    file_texts = []
                    for f in m["files"]:
                        if not f.get("continuous") and i != len(curr_msgs) - 1:
                            continue
                        file_texts.append(f"--- 附件: {f['filename']} ---\n{f['content']}")
                    if file_texts:
                        content = "【参考附件】\n" + "\n".join(file_texts) + "\n\n【用户指令】\n" + content
                api_msgs.append({"role": m["role"], "content": content})

            client, profile = get_client()

            st.session_state.stop_req = False
            st.session_state.is_streaming = True
            st.session_state._stream_content = ""
            st.session_state._stream_chat_id = st.session_state.current_chat_id

            stop_placeholder = st.empty()
            with stop_placeholder.container():
                st.button("⏹️ 停止生成", key="stop_gen", on_click=lambda: st.session_state.update(stop_req=True, is_streaming=False))

            message_placeholder = st.empty()
            full_resp = ""

            try:
                resp = client.chat.completions.create(**build_api_kwargs(profile, api_msgs))
                for chunk in resp:
                    if st.session_state.get("stop_req"):
                        break
                    if chunk.choices and chunk.choices[0].delta.content is not None:
                        full_resp += chunk.choices[0].delta.content
                        st.session_state._stream_content = full_resp
                        message_placeholder.markdown(full_resp + "▌")
            except Exception as e:
                st.error(f"请求失败: {str(e)}")
            finally:
                stop_placeholder.empty()
                is_half = st.session_state.get("stop_req", False)
                if full_resp:
                    last_msg = curr_chat["messages"][-1] if curr_chat["messages"] else {}
                    if not (last_msg.get("role") == "assistant" and last_msg.get("_is_half") and last_msg.get("content") == full_resp):
                        curr_chat["messages"].append({"role": "assistant", "content": full_resp, "_is_half": is_half})

                st.session_state.is_streaming = False
                st.session_state._stream_content = ""
                st.session_state._stream_chat_id = None
                st.session_state.stop_req = False
                trigger_save()
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
                "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 4096,
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
            p["use_max_tokens"] = st.checkbox("📏 Max Tokens", p.get("use_max_tokens", True))
            if p["use_max_tokens"]:
                # 上限放开到 200 万，步长 256，覆盖所有主流模型的最大上下文
                current_mt = p.get("max_tokens", 4096)
                if current_mt > 2000000: current_mt = 2000000
                p["max_tokens"] = st.number_input(
                    "最大 Token 数",
                    min_value=1, max_value=2000000,
                    value=current_mt, step=256,
                    help=(
                        "支持 1 – 2,000,000。\n\n"
                        "• 常见输出上限：GPT-4o ≈ 16K、Claude 3.5 ≈ 8K、Claude 3.7 ≈ 64K、DeepSeek ≈ 8K\n"
                        "• 大上下文窗口：Claude 200K / GPT-4.1 1M / Gemini 1.5-2.0 Pro 2M\n"
                        "• 若接口把此参数当作上下文窗口而非输出上限，可调到相应大小"
                    )
                )
                # 快捷预设（小字按钮，不改变整体布局）
                preset_cols = st.columns(8)
                presets = [("4K", 4096), ("8K", 8192), ("16K", 16384), ("32K", 32768),
                           ("64K", 65536), ("128K", 131072), ("1M", 1048576), ("2M", 2000000)]
                for _ci, (lbl, val) in enumerate(presets):
                    if preset_cols[_ci].button(lbl, key=f"mt_preset_{lbl}", use_container_width=True):
                        p["max_tokens"] = val
                        trigger_save()
                        st.rerun()
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
