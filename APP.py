import streamlit as st
from openai import OpenAI
import io, json, re, requests, uuid, copy, html, time
from datetime import datetime
from docx import Document
from pypdf import PdfReader
from streamlit_local_storage import LocalStorage

# ==========================================
# 1. 页面全局配置与前端美化 (包含布局黑魔法)
# ==========================================
st.set_page_config(page_title="ZenMux 创作者工作站", page_icon="🐙", layout="wide")
st.markdown("""
    <style>
    .stButton>button { border-radius: 8px; font-weight: bold; transition: all 0.3s; }
    button[title="View fullscreen"] {display: none;}
    .css-1jc7ptx, .e1ewe7hr3, .viewerBadge_container__1QSob, .styles_viewerBadge__1yB5_ {display: none;}
    
    /* 手机端极简适配优化 */
    @media (max-width: 768px) {
        .block-container { padding-top: 1rem; padding-bottom: 5rem; }
    }

    /* 🔥 终极黑魔法：将加号按钮强行拽入输入框最左侧 */
    /* 给文本输入区让出左侧的 45px 空间 */
    [data-testid="stChatInput"] textarea {
        padding-left: 3rem !important;
    }
    
    /* 利用 help 属性精准狙击挂载按钮，将其绝对定位到屏幕底部 */
    button[title="挂载附件"] {
        position: fixed !important;
        bottom: 3.2rem; /* 电脑端输入框高度适配 */
        z-index: 99999;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        font-size: 1.2rem !important;
        padding: 0 !important;
        width: 35px !important;
        height: 35px !important;
        transform: translateX(10px); /* 微调位置刚好卡进输入框 */
        color: #666;
    }
    button[title="挂载附件"]:hover {
        color: #000;
        background: #f0f2f5 !important;
        border-radius: 50%;
    }
    
    /* 手机端由于输入框变窄，微调高度 */
    @media (max-width: 768px) {
        button[title="挂载附件"] {
            bottom: 2.1rem;
        }
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 数据持久化层 (移动端防弹版)
# ==========================================
try:
    localS = LocalStorage()
except Exception:
    localS = None

def trigger_save():
    st.session_state._needs_save = True

def execute_save():
    if st.session_state.get("_needs_save", False) and not st.session_state.get("is_streaming", False):
        data = {
            "profiles": st.session_state.profiles,
            "free_chats": st.session_state.free_chats
        }
        if localS:
            try:
                localS.setItem("zenmux_data", json.dumps(data))
            except Exception:
                pass
        st.session_state._needs_save = False

# 初始化状态
if "initialized" not in st.session_state:
    st.session_state.initialized = False
    st.session_state.ls_loaded = False
    st.session_state._needs_save = False
    st.session_state.is_streaming = False
    st.session_state.ls_wait_count = 0

# 水合逻辑 (带手机端拦截逃逸机制)
if not st.session_state.ls_loaded:
    saved_data = None
    if localS:
        try:
            saved_data = localS.getItem("zenmux_data")
        except Exception:
            pass
            
    if saved_data in [None, "", "null"] and st.session_state.ls_wait_count < 1:
        st.session_state.ls_wait_count += 1
        st.info("🔄 正在安全环境中初始化您的专属工作站，请稍候...")
        time.sleep(0.8)
        st.rerun()
        
    default_profiles = [{
        "name": "默认引擎", "base_url": "", "api_key": "", "model": "anthropic/claude-3-5-sonnet-20240620",
        "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 4096,
        "use_top_p": False, "top_p": 1.0, "use_frequency_penalty": False, "frequency_penalty": 0.0
    }]
    first_id = str(uuid.uuid4())
    default_chats = {first_id: {"title": "新对话", "messages": [], "session_knowledge": [], "system_prompt": "", "is_pinned": False, "is_archived": False}}
    
    if saved_data and saved_data not in ["", "null"]:
        try:
            data = json.loads(saved_data) if isinstance(saved_data, str) else saved_data
            if not data or not isinstance(data, dict):
                raise ValueError
            st.session_state.profiles = data.get("profiles", default_profiles)
            st.session_state.free_chats = data.get("free_chats", default_chats)
        except Exception:
            st.session_state.profiles = default_profiles
            st.session_state.free_chats = default_chats
    else:
        st.session_state.profiles = default_profiles
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

def export_to_pretty_html(messages, title, meta=None):
    if meta is None: meta = {}
    css = """
    * { margin:0; padding:0; box-sizing:border-box; }
    body { font-family: -apple-system, sans-serif; background:#f0f2f5; color:#1a1a1a; }
    .header { background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); color:#fff; padding:24px 32px; position:sticky; top:0; z-index:100; box-shadow:0 2px 12px rgba(0,0,0,.15); }
    .header h1 { font-size:22px; font-weight:700; }
    .header .meta { font-size:12px; opacity:.75; margin-top:6px; }
    .chat-container { max-width:860px; margin:0 auto; padding:24px 16px 80px; }
    .info-card { background:#fff; border-radius:12px; padding:20px 24px; margin-bottom:24px; box-shadow:0 1px 4px rgba(0,0,0,.06); border-left:4px solid #667eea; }
    .info-card h3 { font-size:14px; color:#667eea; margin-bottom:12px; font-weight:700; }
    .info-row { display:flex; margin-bottom:8px; font-size:13px; line-height:1.6; }
    .info-label { color:#888; min-width:90px; flex-shrink:0; font-weight:600; }
    .info-value { color:#333; word-break:break-all; }
    .info-value.prompt { background:#f8f8f8; padding:8px 12px; border-radius:6px; font-size:12px; line-height:1.7; margin-top:4px; white-space:pre-wrap; max-height:200px; overflow-y:auto; }
    .msg { display:flex; gap:12px; margin-bottom:24px; align-items:flex-start; }
    .msg.user { flex-direction:row-reverse; }
    .avatar { width:36px; height:36px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:18px; flex-shrink:0; }
    .msg.ai .avatar { background:#e8f5e9; }
    .msg.user .avatar { background:#e3f2fd; }
    .bubble { max-width:75%; padding:14px 18px; border-radius:16px; line-height:1.8; font-size:15px; word-wrap:break-word; white-space:pre-wrap; box-shadow:0 1px 3px rgba(0,0,0,.06); }
    .msg.ai .bubble { background:#fff; border-top-left-radius:4px; }
    .msg.user .bubble { background:#d1e7ff; border-top-right-radius:4px; }
    .word-count { font-size:11px; color:#999; margin-top:6px; text-align:right; }
    .msg.user .word-count { text-align:left; }
    """
    js = "<script>function toggleInfo() { var el = document.getElementById('infoCard'); el.style.display = (el.style.display==='none') ? 'block' : 'none'; }</script>"
    
    info_html = ""
    if any(meta.get(k) for k in ["system_prompt", "model", "files"]):
        rows = ""
        if meta.get("model"): rows += f'<div class="info-row"><span class="info-label">🧠 模型</span><span class="info-value">{meta["model"]}</span></div>'
        if meta.get("system_prompt"):
            rows += f'<div class="info-row"><span class="info-label">🎭 人设</span></div><div class="info-value prompt">{html.escape(meta["system_prompt"])}</div>'
        info_html = f'<div class="info-card" id="infoCard"><h3>⚙️ 对话配置信息</h3>{rows}</div>'

    msg_html = ""
    for m in messages:
        if m["role"] == "system": continue
        is_user = m["role"] == "user"
        role_class = "user" if is_user else "ai"
        avatar = "🙋‍♂️" if is_user else "🤖"
        safe = html.escape(m["content"]).replace('\n', '<br>')
        msg_html += f'<div class="msg {role_class}"><div class="avatar">{avatar}</div><div><div class="bubble">{safe}</div><div class="word-count">{count_words(m["content"])} 字</div></div></div>'

    date_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    return f"<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>{title}</title><style>{css}</style></head><body><div class='header'><h1>💬 {title}</h1><div class='meta'>{date_str} <span style='text-decoration:underline;cursor:pointer;margin-left:10px' onclick='toggleInfo()'>显示/隐藏配置</span></div></div><div class='chat-container'>{info_html}{msg_html}</div>{js}</body></html>".encode('utf-8')

def fetch_models(base_url, api_key):
    try:
        url = (base_url.strip().rstrip('/') or "[https://api.openai.com/v1](https://api.openai.com/v1)") + "/models"
        resp = requests.get(url, headers={"Authorization": "Bearer " + api_key.strip()}, timeout=8)
        if resp.status_code == 200:
            return True, sorted([m["id"] for m in resp.json().get("data", [])])
        return False, f"状态码 {resp.status_code}: {resp.text[:100]}"
    except Exception as e:
        return False, str(e)

def get_client():
    p = st.session_state.profiles[st.session_state.active_profile_idx]
    url = p["base_url"].strip() or "[https://api.openai.com/v1](https://api.openai.com/v1)"
    return OpenAI(base_url=url, api_key=p["api_key"].strip()), p

def build_api_kwargs(profile, api_msgs):
    kw = {"model": profile["model"], "messages": api_msgs, "stream": True}
    if profile.get("use_temperature", True): kw["temperature"] = profile.get("temperature", 0.8)
    if profile.get("use_max_tokens", True): kw["max_tokens"] = profile.get("max_tokens", 4096)
    if profile.get("use_top_p", False): kw["top_p"] = profile.get("top_p", 1.0)
    if profile.get("use_frequency_penalty", False): kw["frequency_penalty"] = profile.get("frequency_penalty", 0.0)
    return kw

# ==========================================
# 4. 全局侧边栏导航 (已整合会话列表)
# ==========================================
with st.sidebar:
    st.header("控制中枢")
    pages = ["💬 自由聊天区", "⚙️ 底层引擎配置"]
    for pg in pages:
        btype = "primary" if st.session_state.current_page == pg else "secondary"
        if st.button(pg, use_container_width=True, type=btype):
            st.session_state.current_page = pg
            st.rerun()
            
    # 🌟 优化：将会话列表彻底移入左侧边栏
    if st.session_state.current_page == "💬 自由聊天区":
        st.divider()
        st.markdown("### 📚 会话管理")
        col_new, col_search = st.columns([1, 2])
        if col_new.button("➕ 新对话", use_container_width=True, type="primary"):
            nid = str(uuid.uuid4())
            st.session_state.free_chats[nid] = {"title": "新对话", "messages": [], "session_knowledge": [], "system_prompt": "", "is_pinned": False, "is_archived": False}
            st.session_state.current_chat_id = nid
            trigger_save()
            st.rerun()
        search_q = col_search.text_input("🔍 搜索历史", label_visibility="collapsed", placeholder="搜索历史对话...")

        chat_items = [(cid, cdata) for cid, cdata in st.session_state.free_chats.items() if not cdata.get("is_archived", False) and (not search_q or search_q.lower() in cdata["title"].lower())]
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
            st.caption("🗄️ 归档会话")
            for cid, cdata in st.session_state.free_chats.items():
                if cdata.get("is_archived", False):
                    cc1, cc2 = st.columns([3, 1])
                    cc1.markdown(f"📦 {cdata['title']}")
                    if cc2.button("恢复", key=f"unarch_{cid}"):
                        cdata["is_archived"] = False
                        trigger_save()
                        st.rerun()
            
    active_p = st.session_state.profiles[st.session_state.active_profile_idx]
    st.divider()
    st.caption(f"🟢 **当前挂载**: {active_p['name']}\n🧠 **模型**: {active_p['model']}")

    st.divider()
    with st.expander("📦 全量资产导出恢复舱", expanded=False):
        st.caption("换设备时一键导入导出所有数据。")
        full_data = json.dumps({"profiles": st.session_state.profiles, "free_chats": st.session_state.free_chats}, ensure_ascii=False, indent=2).encode('utf-8')
        st.download_button("📥 导出全量快照包", full_data, f"ZenMux_Backup_{datetime.now().strftime('%m%d_%H%M')}.json", "application/json", use_container_width=True, type="primary")
        
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
# 模块 1: 自由聊天区 (极致清爽版)
# ==========================================
if st.session_state.current_page == "💬 自由聊天区":
    # --- 当前会话主体 ---
    if st.session_state.current_chat_id not in st.session_state.free_chats:
        st.session_state.current_chat_id = list(st.session_state.free_chats.keys())[-1]
    curr_chat = st.session_state.free_chats[st.session_state.current_chat_id]
    
    # 标题与终极聚合操作菜单
    tc1, tc2 = st.columns([15, 1]) 
    with tc1:
        new_title = st.text_input("会话标题", curr_chat["title"], label_visibility="collapsed")
        if new_title != curr_chat["title"]:
            curr_chat["title"] = new_title
            trigger_save()
    with tc2:
        with st.popover("🔽", use_container_width=True):
            st.markdown("##### ⚙️ 会话管理")
            btn_c1, btn_c2 = st.columns(2)
            if btn_c1.button("取消置顶" if curr_chat.get("is_pinned") else "📌 置顶", use_container_width=True):
                curr_chat["is_pinned"] = not curr_chat.get("is_pinned", False)
                trigger_save()
                st.rerun()
            if btn_c2.button("📦 归档", use_container_width=True):
                curr_chat["is_archived"] = True
                trigger_save()
                st.rerun()
            if btn_c1.button("🗑️ 清空", use_container_width=True):
                curr_chat["messages"] = []
                trigger_save()
                st.rerun()
            if btn_c2.button("📥 导出", use_container_width=True):
                st.session_state._show_export = not st.session_state.get("_show_export", False)
                st.rerun()
                
            st.divider()
            
            st.markdown("##### 📚 全局设定与知识库")
            curr_chat["system_prompt"] = st.text_area("🎭 System Prompt", curr_chat.get("system_prompt", ""), height=80, placeholder="设定此会话专属的全局人设...")
            up_f = st.file_uploader("📎 上传常驻参考文件", type=['txt', 'md', 'pdf', 'docx'], key=f"kb_{st.session_state.current_chat_id}")
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

    if st.session_state.get("_show_export", False):
        with st.container(border=True):
            st.markdown("**📦 导出当前对话**")
            exp_mode = st.radio("格式", ["完整记录", "纯享正文"], horizontal=True, label_visibility="collapsed")
            is_pure = (exp_mode == "纯享正文")
            curr_msgs = curr_chat["messages"]
            
            txt_c = "\n\n".join([clean_novel_text(m['content']) for m in curr_msgs if m['role'] == 'assistant']) if is_pure else "\n".join([f"{'我' if m['role']=='user' else 'AI'}:\n{m['content']}\n\n{'-'*40}\n" for m in curr_msgs])
            st.download_button("📥 下载 TXT", txt_c.encode('utf-8'), f"{curr_chat['title']}.txt", use_container_width=True)
            st.download_button("📥 下载 Word", generate_word_doc(curr_msgs, is_pure), f"{curr_chat['title']}.docx", use_container_width=True)
            export_meta = {"system_prompt": curr_chat.get("system_prompt", ""), "model": active_p["model"]}
            st.download_button("🎨 下载 HTML", export_to_pretty_html(curr_msgs, curr_chat["title"], export_meta), f"{curr_chat['title']}.html", "text/html", use_container_width=True)

    # 聊天消息展示区
    with st.container(border=False):
        editing_idx = st.session_state.get("_editing_chat_idx")
        
        for i, msg in enumerate(curr_chat["messages"]):
            if msg["role"] == "system": continue
            
            with st.chat_message(msg["role"]):
                if msg.get("files"):
                    for f in msg["files"]:
                        st.caption(f"`{'🔄' if f.get('continuous') else '1️⃣'} 附件: {f['filename']}`")
                        
                if msg["role"] == "user" and editing_idx == i:
                    new_text = st.text_area("✏️ 编辑", msg["content"], key=f"edit_area_{i}", height=100)
                    ebc1, ebc2 = st.columns(2)
                    if ebc1.button("✅ 重新发送", key=f"edit_ok_{i}", type="primary"):
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
                    
                    if msg["role"] == "user" and editing_idx is None:
                        with st.popover("⚙️ 操作"):
                            if st.button("✏️ 编辑", key=f"edit_btn_{i}"):
                                st.session_state._editing_chat_idx = i
                                st.rerun()
                            if st.button("🗑️ 删除此条", key=f"del_msg_{i}"):
                                curr_chat["messages"].pop(i)
                                trigger_save()
                                st.rerun()
                    elif msg["role"] == "assistant":
                        with st.popover("⚙️ 操作"):
                            if st.button("🔄 重新生成", key=f"regen_{i}"):
                                curr_chat["messages"] = curr_chat["messages"][:i]
                                st.session_state._auto_resend = True
                                trigger_save()
                                st.rerun()
                            if msg.get("_is_half"):
                                if st.button("▶️ 续写", key=f"resume_{i}"):
                                    st.session_state._resume_idx = i
                                    st.session_state._auto_resend = True
                                    st.rerun()
                            if st.button("🗑️ 删除此条", key=f"del_msg_{i}"):
                                curr_chat["messages"].pop(i)
                                trigger_save()
                                st.rerun()
                        st.caption(f"📊 {count_words(msg['content'])} 字")

    # --- 输入区 ---
    need_resend = st.session_state.pop("_auto_resend", False)
    resume_idx = st.session_state.pop("_resume_idx", None)
    
    # 🌟 优化：挂载按钮。利用顶部的 CSS 黑魔法，它会自动跑到输入框内部的最左侧去！
    with st.popover("➕", help="挂载附件"):
        dyn_file = st.file_uploader("📎 上传单次文档", type=['txt', 'md', 'pdf', 'docx'], key="dyn_file")
        is_continuous = st.checkbox("🔄 持续参考 (勾选后一直带入后续对话)", value=False)

    prompt = st.chat_input("输入消息...")

    if prompt or need_resend:
        if not active_p["api_key"]:
            st.error("⚠️ 请先在【底层引擎配置】中填写 API Key！")
            st.stop()

        if resume_idx is not None:
            prompt = "请紧接上文最后一个字继续往下写，保持文风和节奏一致。"
            curr_chat["messages"][resume_idx]["_is_half"] = False
            
        if prompt:
            if len(curr_chat["messages"]) == 0 and curr_chat["title"] == "新对话":
                curr_chat["title"] = prompt[:10] + ("..." if len(prompt)>10 else "")
            new_msg = {"role": "user", "content": prompt}
            if dyn_file and not need_resend:
                new_msg["files"] = [{"filename": dyn_file.name, "content": extract_file_text(dyn_file), "continuous": is_continuous}]
            curr_chat["messages"].append(new_msg)
            trigger_save()

        if prompt:
            with st.chat_message("user"):
                st.markdown(prompt)

        api_msgs = []
        sp = curr_chat.get("system_prompt", "").strip()
        if sp: api_msgs.append({"role": "system", "content": sp})
        if curr_chat.get("session_knowledge"):
            kb_parts = [f"--- {k['filename']} ---\n{k['content']}" for k in curr_chat["session_knowledge"]]
            api_msgs.append({"role": "system", "content": "【全局参考】：\n" + "\n\n".join(kb_parts)})

        for i, m in enumerate(curr_chat["messages"]):
            content = m["content"]
            if m.get("files"):
                file_texts = [f"--- {f['filename']} ---\n{f['content']}" for f in m["files"] if f.get("continuous") or i == len(curr_chat["messages"]) - 1]
                if file_texts:
                    content = "【附件】\n" + "\n".join(file_texts) + "\n\n【指令】\n" + content
            api_msgs.append({"role": m["role"], "content": content})

        client, profile = get_client()
        with st.chat_message("assistant"):
            stop_btn = st.button("⏹️ 停止生成", key="stop_gen")
            message_placeholder = st.empty()
            full_resp = ""
            st.session_state.is_streaming = True
            try:
                resp = client.chat.completions.create(**build_api_kwargs(profile, api_msgs))
                for chunk in resp:
                    if stop_btn: break
                    if chunk.choices and chunk.choices[0].delta.content is not None:
                        full_resp += chunk.choices[0].delta.content
                        message_placeholder.markdown(full_resp + "▌")
                message_placeholder.markdown(full_resp)
                curr_chat["messages"].append({"role": "assistant", "content": full_resp, "_is_half": bool(stop_btn)})
            except Exception as e:
                st.error(f"请求失败: {str(e)}")
                if full_resp: curr_chat["messages"].append({"role": "assistant", "content": full_resp, "_is_half": True})
            finally:
                st.session_state.is_streaming = False
                trigger_save()
                st.rerun()

# ==========================================
# 模块 2: 底层引擎配置
# ==========================================
elif st.session_state.current_page == "⚙️ 底层引擎配置":
    st.header("⚙️ 底层驱动配置")
    
    p_names = [p["name"] for p in st.session_state.profiles]
    idx = st.radio("切换引擎", range(len(p_names)), format_func=lambda x: p_names[x], index=st.session_state.active_profile_idx, horizontal=True)
    st.session_state.active_profile_idx = idx
    
    if st.button("➕ 新增引擎", use_container_width=True):
        st.session_state.profiles.append({
            "name": f"新引擎 {len(p_names) + 1}", "base_url": "", "api_key": "", "model": "anthropic/claude-3-5-sonnet-20240620",
            "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 4096,
            "use_top_p": False, "top_p": 1.0, "use_frequency_penalty": False, "frequency_penalty": 0.0
        })
        trigger_save()
        st.rerun()

    st.divider()
    p = st.session_state.profiles[idx]

    c1, c2 = st.columns([3, 1])
    p["name"] = c1.text_input("引擎标签", p["name"])
    if c2.button("💾 保存配置", type="primary", use_container_width=True):
        trigger_save()
        st.success("已保存！")

    p["base_url"] = st.text_input("Base URL", p["base_url"])
    p["api_key"] = st.text_input("API Key", p["api_key"], type="password")
    st.caption("🔒 你的 Key 只保存在浏览器本地缓存中，服务器不会记录。")

    if p["api_key"] and st.button("🔑 测试连通性"):
        with st.spinner("测试中..."):
            try:
                OpenAI(base_url=p["base_url"].strip() or "[https://api.openai.com/v1](https://api.openai.com/v1)", api_key=p["api_key"].strip()).chat.completions.create(model=p["model"], messages=[{"role": "user", "content": "Hi"}], max_tokens=5)
                st.success("✅ 连通成功！")
            except Exception as e:
                st.error(f"❌ 失败: {str(e)}")

    m1, m2 = st.columns([3, 1])
    p["model"] = m1.text_input("模型映射 (Model ID)", p["model"])
    if m2.button("🔄 获取列表"):
        if p["api_key"]:
            with st.spinner("获取中..."):
                success, result = fetch_models(p["base_url"], p["api_key"])
                if success and result:
                    st.session_state.temp_models = result
                    st.success(f"✅ 获取到 {len(result)} 个模型！")
                else:
                    st.error(f"❌ 失败: {result}")

    if "temp_models" in st.session_state:
        sel_m = st.selectbox("选择模型", ["(不覆盖)"] + st.session_state.temp_models)
        if sel_m != "(不覆盖)":
            p["model"] = sel_m
            del st.session_state.temp_models
            trigger_save()
            st.rerun()

    with st.expander("🎛️ 运行时超参数"):
        p["use_temperature"] = st.checkbox("🔥 Temperature", p.get("use_temperature", True))
        if p["use_temperature"]: p["temperature"] = st.slider("温度", 0.0, 2.0, p.get("temperature", 0.8), 0.1)
        p["use_max_tokens"] = st.checkbox("📏 Max Tokens", p.get("use_max_tokens", True))
        if p["use_max_tokens"]: p["max_tokens"] = st.slider("最大Token", 512, 16384, p.get("max_tokens", 4096), 512)

    if len(st.session_state.profiles) > 1 and st.button("🗑️ 删除此引擎", type="primary"):
        st.session_state.profiles.pop(idx)
        st.session_state.active_profile_idx = 0
        trigger_save()
        st.rerun()

execute_save()
