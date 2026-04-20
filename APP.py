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
    /* 手机端适配优化 */
    @media (max-width: 768px) {
        .block-container { padding-top: 2rem; padding-bottom: 5rem; }
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 数据持久化层 (LocalStorage 桥接)
# ==========================================
localS = LocalStorage()

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
    st.session_state.ls_wait_count = 0  # 等待计数器，防止新用户死锁

# 水合逻辑 (从 LocalStorage 读取)
if not st.session_state.ls_loaded:
    saved_data = localS.getItem("zenmux_data", key="ls_get")
    st.session_state.ls_wait_count += 1
    
    # 如果拿到了数据，或者已经等了1个周期（说明是新用户，本地没数据），就放行
    if saved_data is not None or st.session_state.ls_wait_count > 1: 
        default_profiles = [{
            "name": "默认引擎", "base_url": "", "api_key": "", "model": "anthropic/claude-3-5-sonnet-20240620",
            "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 4096,
            "use_top_p": False, "top_p": 1.0, "use_frequency_penalty": False, "frequency_penalty": 0.0
        }]
        first_id = str(uuid.uuid4())
        default_chats = {first_id: {"title": "新对话", "messages": [], "session_knowledge": [], "system_prompt": "", "is_pinned": False, "is_archived": False}}
        
        if saved_data:
            try:
                data = json.loads(saved_data) if isinstance(saved_data, str) else saved_data
                st.session_state.profiles = data.get("profiles", default_profiles)
                st.session_state.free_chats = data.get("free_chats", default_chats)
            except:
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
    else:
        st.info("🔄 正在从本地安全存储加载数据，请稍候...")
        st.stop()

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
    body { font-family: -apple-system, 'PingFang SC', 'Microsoft YaHei', sans-serif; background:#f0f2f5; color:#1a1a1a; }
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
    .file-tag { display:inline-block; background:#f0f0f0; padding:2px 10px; border-radius:12px; font-size:12px; color:#555; margin:2px 4px 2px 0; }
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
    .footer { text-align:center; padding:32px; font-size:12px; color:#aaa; border-top:1px solid #e5e5e5; max-width:860px; margin:0 auto; }
    """
    js = """
    <script>
    function toggleInfo() {
        var el = document.getElementById('infoCard');
        if(el.style.display==='none'){el.style.display='block';}else{el.style.display='none';}
    }
    </script>
    """
    info_html = ""
    has_meta = any(meta.get(k) for k in ["source", "system_prompt", "model", "files"])
    if has_meta:
        rows = ""
        if meta.get("model"): rows += f'<div class="info-row"><span class="info-label">🧠 模型</span><span class="info-value">{meta["model"]}</span></div>'
        if meta.get("files"):
            tags = "".join([f'<span class="file-tag">📄 {f["filename"]} ({f.get("size", 0):,} 字)</span>' for f in meta["files"]])
            rows += f'<div class="info-row"><span class="info-label">📎 挂载文件</span><span class="info-value">{tags}</span></div>'
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
        wc = count_words(m["content"])
        msg_html += f'<div class="msg {role_class}"><div class="avatar">{avatar}</div><div><div class="bubble">{safe}</div><div class="word-count">{wc} 字</div></div></div>'

    date_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    total_ai_words = sum(count_words(m["content"]) for m in messages if m["role"] == "assistant")
    toggle_link = '<span style="margin-left:12px;text-decoration:underline;cursor:pointer;" onclick="toggleInfo()">展开/收起配置</span>' if has_meta else ""
    
    full_html = f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{title}</title><style>{css}</style></head><body><div class="header"><h1>💬 {title}</h1><div class="meta">{date_str} | AI 共 {total_ai_words:,} 字 {toggle_link}</div></div><div class="chat-container">{info_html}{msg_html}</div><div class="footer">ZenMux AI 创作者工作站</div>{js}</body></html>"""
    return full_html.encode('utf-8')

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
    if profile.get("use_max_tokens", True): kw["max_tokens"] = profile.get("max_tokens", 4096)
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

    # 全量资产导出恢复舱
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
    # --- 顶部会话管理 (适配手机端) ---
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

        # 过滤与排序 (置顶优先，排除归档)
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
    tc1, tc2, tc3, tc4, tc5 = st.columns([4, 1, 1, 1, 1])
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
                
            ec1, ec2, ec3 = st.columns(3)
            ec1.download_button("📥 下载 TXT", txt_c.encode('utf-8'), f"{curr_chat['title']}.txt", use_container_width=True)
            ec2.download_button("📥 下载 Word", generate_word_doc(curr_msgs, is_pure), f"{curr_chat['title']}.docx", use_container_width=True)
            
            # 构造 HTML 导出的元数据
            export_meta = {
                "system_prompt": curr_chat.get("system_prompt", ""),
                "model": active_p["model"],
                "files": [{"filename": k["filename"], "size": len(k["content"])} for k in curr_chat.get("session_knowledge", [])]
            }
            ec3.download_button("🎨 下载 HTML", export_to_pretty_html(curr_msgs, curr_chat["title"], export_meta), f"{curr_chat['title']}.html", "text/html", use_container_width=True)

    # --- 对话设置区 (常驻知识库) ---
    has_content = bool(curr_chat.get("session_knowledge") or curr_chat.get("system_prompt"))
    with st.expander("⚙️ 全局设定与常驻知识库", expanded=has_content):
        curr_chat["system_prompt"] = st.text_area("🎭 System Prompt (全局人设)", curr_chat.get("system_prompt", ""), height=80)
        up_f = st.file_uploader("📎 上传常驻参考文件 (TXT/MD/PDF/DOCX)", type=['txt', 'md', 'pdf', 'docx'], key=f"kb_{st.session_state.current_chat_id}")
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
                # 渲染随消息挂载的文件徽章
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
                    
                    # 操作栏
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
        dyn_file = st.file_uploader("上传文件 (TXT/MD/PDF/DOCX)", type=['txt', 'md', 'pdf', 'docx'], key="dyn_file")
        is_continuous = st.checkbox("🔄 持续参考 (开启后该文件将一直带入后续对话，否则仅本次有效)", value=False)

    prompt = st.chat_input("输入消息...")

    # --- 发送与流式请求逻辑 ---
    if prompt or need_resend:
        if not active_p["api_key"]:
            st.error("⚠️ 请先在左侧【底层引擎配置】中填写 API Key！")
            st.stop()

        # 处理续写逻辑
        if resume_idx is not None:
            prompt = "请紧接上文最后一个字继续往下写，保持文风和节奏一致。"
            curr_chat["messages"][resume_idx]["_is_half"] = False # 移除半成品标记
            
        if prompt:
            # 自动命名标题
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
            trigger_save()

        if prompt:
            with st.chat_message("user"):
                st.markdown(prompt)

        # 构建 API 消息包
        api_msgs = []
        
        # 1. System Prompt
        sp = curr_chat.get("system_prompt", "").strip()
        if sp: api_msgs.append({"role": "system", "content": sp})
        
        # 2. 常驻知识库
        if curr_chat.get("session_knowledge"):
            kb_parts = [f"--- 文件: {k['filename']} ---\n{k['content']}" for k in curr_chat["session_knowledge"]]
            api_msgs.append({"role": "system", "content": "【全局参考文件】：\n" + "\n\n".join(kb_parts)})

        # 3. 历史消息与动态附件
        for i, m in enumerate(curr_chat["messages"]):
            content = m["content"]
            # 处理动态附件
            if m.get("files"):
                file_texts = []
                for f in m["files"]:
                    # 如果是一次性文件，且不是当前最新一轮，则跳过
                    if not f.get("continuous") and i != len(curr_chat["messages"]) - 1:
                        continue
                    file_texts.append(f"--- 附件: {f['filename']} ---\n{f['content']}")
                if file_texts:
                    content = "【参考附件】\n" + "\n".join(file_texts) + "\n\n【用户指令】\n" + content
                    
            api_msgs.append({"role": m["role"], "content": content})

        # 4. 调用 API
        client, profile = get_client()
        with st.chat_message("assistant"):
            stop_btn = st.button("⏹️ 停止生成", key="stop_gen")
            message_placeholder = st.empty()
            full_resp = ""
            st.session_state.is_streaming = True
            
            try:
                resp = client.chat.completions.create(**build_api_kwargs(profile, api_msgs))
                for chunk in resp:
                    if stop_btn: # 检查中断
                        break
                    if chunk.choices and chunk.choices[0].delta.content is not None:
                        full_resp += chunk.choices[0].delta.content
                        message_placeholder.markdown(full_resp + "▌")
                        
                message_placeholder.markdown(full_resp)
                render_copy_button(full_resp)
                
                # 保存结果
                is_half = bool(stop_btn)
                curr_chat["messages"].append({"role": "assistant", "content": full_resp, "_is_half": is_half})
                
            except Exception as e:
                st.error(f"请求失败: {str(e)}")
                if full_resp: # 保留报错前的半成品
                    curr_chat["messages"].append({"role": "assistant", "content": full_resp, "_is_half": True})
            finally:
                st.session_state.is_streaming = False
                trigger_save()
                st.rerun()

# ==========================================
# 模块 2: 底层引擎配置 (原样保留)
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
            if p["use_temperature"]: p["temperature"] = st.slider("温度值", 0.0, 2.0, p.get("temperature", 0.8), 0.1, label_visibility="collapsed")
            p["use_max_tokens"] = st.checkbox("📏 Max Tokens", p.get("use_max_tokens", True))
            if p["use_max_tokens"]: p["max_tokens"] = st.slider("最大Token", 512, 16384, p.get("max_tokens", 4096), 512, label_visibility="collapsed")
        with sl2:
            p["use_top_p"] = st.checkbox("🎲 Top P", p.get("use_top_p", False))
            if p["use_top_p"]: p["top_p"] = st.slider("Top P值", 0.0, 1.0, p.get("top_p", 1.0), 0.05, label_visibility="collapsed")
            p["use_frequency_penalty"] = st.checkbox("🚫 Frequency Penalty", p.get("use_frequency_penalty", False))
            if p["use_frequency_penalty"]: p["frequency_penalty"] = st.slider("惩罚值", -2.0, 2.0, p.get("frequency_penalty", 0.0), 0.1, label_visibility="collapsed")

        if len(st.session_state.profiles) > 1:
            st.divider()
            if st.button("🗑️ 删除此引擎"):
                st.session_state.profiles.pop(idx)
                st.session_state.active_profile_idx = 0
                trigger_save()
                st.rerun()

# 统一执行保存
execute_save()
