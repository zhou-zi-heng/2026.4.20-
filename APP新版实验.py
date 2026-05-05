import streamlit as st
from openai import OpenAI
import io, base64, json, re, requests, uuid, html, time
from datetime import datetime
from docx import Document
from pypdf import PdfReader
from streamlit_local_storage import LocalStorage

# ==========================================
# 1. 页面全局配置与全平台兼容极简 UI
# ==========================================
st.set_page_config(page_title="ZenMux 创作者工作站", page_icon="🐙", layout="wide")
st.markdown("""
    <style>
    [data-testid="collapsedControl"] { display: flex !important; visibility: visible !important; z-index: 999999 !important; }

    .block-container { padding-top: 3.5rem !important; padding-bottom: 6rem !important; }

    div.stMain div[data-testid="stHorizontalBlock"]:first-of-type {
        position: sticky !important;
        top: 2.8rem !important; 
        z-index: 990 !important;
        background-color: var(--background-color, #ffffff) !important;
        padding: 5px 15px !important; 
        margin-top: -15px !important;
        border-bottom: 1px solid #e5e7eb !important;
        align-items: center !important;
        flex-wrap: nowrap !important;
    }
    div.stMain div[data-testid="stHorizontalBlock"]:first-of-type [data-testid="stTextInput"] { margin: 0 !important; padding: 0 !important; }
    div.stMain div[data-testid="stHorizontalBlock"]:first-of-type div[data-baseweb="input"] {
        background-color: transparent !important; border: none !important; box-shadow: none !important;
        min-height: 0 !important; padding: 0 !important; margin: 0 !important;
    }
    div.stMain div[data-testid="stHorizontalBlock"]:first-of-type input {
        font-size: 18px !important; font-weight: bold !important; padding: 0 !important; margin: 0 !important;
        height: auto !important; color: var(--text-color, #1f2937) !important;
    }
    div.stMain div[data-testid="stHorizontalBlock"]:first-of-type [data-testid="stPopover"] { display: flex; align-items: center; justify-content: flex-end; margin: 0 !important; padding: 0 !important; }
    div.stMain div[data-testid="stHorizontalBlock"]:first-of-type [data-testid="stPopover"] > button {
        background: transparent !important; border: none !important; box-shadow: none !important;
        padding: 0 5px !important; margin: 0 !important; height: auto !important; font-size: 20px !important; color: #9ca3af !important;
    }
    div.stMain div[data-testid="stHorizontalBlock"]:first-of-type [data-testid="stPopover"] > button:hover { color: #667eea !important; }

    div.stMain div[data-testid="stPopover"]:last-of-type {
        position: fixed !important;
        bottom: 85px !important; 
        z-index: 99999 !important;
    }
    @media (min-width: 768px) {
        div.stMain div[data-testid="stPopover"]:last-of-type { left: max(20px, calc(50vw - 360px)) !important; }
    }
    @media (max-width: 768px) {
        div.stMain div[data-testid="stPopover"]:last-of-type { bottom: 75px !important; left: 10px !important; }
    }
    div.stMain div[data-testid="stPopover"]:last-of-type > button {
        background-color: var(--background-color, #ffffff) !important; 
        border: 1px solid #d1d5db !important; 
        border-radius: 20px !important;
        padding: 4px 16px !important; 
        box-shadow: 0 4px 10px rgba(0,0,0,0.08) !important; 
        color: var(--text-color, #374151) !important; 
        font-weight: normal !important;
    }
    div.stMain div[data-testid="stPopover"]:last-of-type > button:hover { border-color: #667eea !important; color: #667eea !important; }

    .zm-copy-btn {
        border:none;background:transparent;color:#888;cursor:pointer;font-size:12px;
        padding:2px 8px;border-radius:6px;margin-left:6px;
    }
    .zm-copy-btn:hover { background:rgba(102,126,234,0.1); color:#667eea; }
    </style>
    <script>
    window.zmCopy = function(btn, b64) {
        try {
            const text = decodeURIComponent(escape(atob(b64)));
            navigator.clipboard.writeText(text).then(function(){
                const old = btn.innerText;
                btn.innerText = '✅ 已复制';
                setTimeout(function(){ btn.innerText = old; }, 1800);
            });
        } catch(e) { alert('复制失败：' + e); }
    };
    </script>
""", unsafe_allow_html=True)

# ==========================================
# 2. 常量与数据持久化层
# ==========================================
SUPPORTED_FILE_TYPES = ['txt', 'md', 'py', 'json', 'csv', 'log', 'pdf', 'docx']
FILE_TYPE_HINT = "TXT/MD/PY/JSON/CSV/LOG/PDF/DOCX"
DEFAULT_RENDER_WINDOW = 50
STORAGE_KEY = "zenmux_data_v2"
SAVE_THROTTLE_SEC = 2

# 白名单字段 - 防止异常对象混入导致 JSON 失败
PROFILE_FIELDS = {"name", "base_url", "api_key", "model",
                  "use_temperature", "temperature",
                  "use_max_tokens", "max_tokens",
                  "use_top_p", "top_p",
                  "use_frequency_penalty", "frequency_penalty"}
CHAT_FIELDS = {"title", "messages", "session_knowledge", "system_prompt",
               "is_pinned", "is_archived"}
MESSAGE_FIELDS = {"role", "content", "files", "_is_half"}

# LocalStorage 容量上限（浏览器通常 5-10MB）
STORAGE_WARN_BYTES = 4 * 1024 * 1024   # 4MB 警告
STORAGE_LIMIT_BYTES = 9 * 1024 * 1024  # 9MB 拒绝

@st.cache_resource
def get_local_storage():
    try:
        return LocalStorage()
    except Exception as e:
        print(f"[ZenMux] LocalStorage 初始化失败: {e}")
        return None

localS = get_local_storage()

def _safe_toast(msg, icon=None):
    """兼容无 st.toast 的旧版本"""
    try:
        if hasattr(st, "toast"):
            st.toast(msg, icon=icon)
        else:
            st.warning(msg)
    except Exception:
        pass

def _sanitize_for_save():
    """清洗数据，只保留白名单字段，避免 JSON 序列化异常"""
    clean_profiles = []
    for p in st.session_state.profiles:
        if not isinstance(p, dict):
            continue
        cp = {k: v for k, v in p.items() if k in PROFILE_FIELDS}
        clean_profiles.append(cp)

    clean_chats = {}
    for cid, c in st.session_state.free_chats.items():
        if not isinstance(c, dict):
            continue
        cc = {k: v for k, v in c.items() if k in CHAT_FIELDS}
        # 清洗 messages
        if "messages" in cc and isinstance(cc["messages"], list):
            clean_msgs = []
            for m in cc["messages"]:
                if isinstance(m, dict):
                    cm = {k: v for k, v in m.items() if k in MESSAGE_FIELDS}
                    clean_msgs.append(cm)
            cc["messages"] = clean_msgs
        clean_chats[cid] = cc

    return clean_profiles, clean_chats

def trigger_save():
    st.session_state._needs_save = True

def execute_save():
    """
    修复 3.2/3.4/3.5：
    - 流式中不保存
    - 节流 2 秒
    - 白名单清洗
    - 容量检查
    - 异常可见
    """
    if not st.session_state.get("_needs_save", False):
        return
    if st.session_state.get("is_streaming", False):
        return

    now = time.time()
    last_save = st.session_state.get("_last_save_ts", 0)
    if now - last_save < SAVE_THROTTLE_SEC:
        return

    if localS is None:
        st.session_state._needs_save = False
        return

    try:
        clean_profiles, clean_chats = _sanitize_for_save()
        data = {
            "profiles": clean_profiles,
            "free_chats": clean_chats,
            "_v": 2,
            "_ts": int(now),
        }
        payload = json.dumps(data, ensure_ascii=False)
        payload_size = len(payload.encode('utf-8'))

        # 容量监控
        st.session_state._last_payload_size = payload_size

        if payload_size > STORAGE_LIMIT_BYTES:
            _safe_toast(f"⚠️ 数据体积已达 {payload_size/1024/1024:.1f}MB，超出浏览器限制，保存已取消。请归档或清理旧对话！", icon="🚫")
            st.session_state._needs_save = False
            return

        if payload_size > STORAGE_WARN_BYTES:
            _safe_toast(f"💾 存储已占用 {payload_size/1024/1024:.1f}MB，建议及时导出快照备份。", icon="⚠️")

        localS.setItem(STORAGE_KEY, payload, key="zm_storage")
        st.session_state._last_save_ts = now
        st.session_state._needs_save = False
    except Exception as e:
        print(f"[ZenMux] 保存失败: {e}")
        _safe_toast(f"保存失败: {str(e)[:80]}", icon="⚠️")
        st.session_state._needs_save = False

# ==========================================
# 导出模态框
# ==========================================
dialog_decorator = getattr(st, "dialog", getattr(st, "experimental_dialog", None))
if dialog_decorator:
    @dialog_decorator("📦 导出对话记录")
    def render_export_modal(curr_chat, active_p):
        st.write("请选择您需要的导出格式：")
        exp_mode = st.radio("导出格式", ["完整记录 (含您的提问)", "纯享正文 (仅提取 AI 回答)"], horizontal=True, label_visibility="collapsed")
        is_pure = (exp_mode == "纯享正文 (仅提取 AI 回答)")
        txt_c = "\n\n".join([clean_novel_text(m['content']) for m in curr_chat["messages"] if m['role'] == 'assistant']) if is_pure else "\n".join([f"{'我' if m['role']=='user' else 'AI'}:\n{m['content']}\n\n{'-'*40}\n" for m in curr_chat["messages"]])

        c1, c2, c3 = st.columns(3)
        c1.download_button("📥 存为 TXT", txt_c.encode('utf-8'), f"{curr_chat['title']}.txt", mime="text/plain", use_container_width=True)
        c2.download_button("📥 存为 Word", generate_word_doc(curr_chat["messages"], is_pure), f"{curr_chat['title']}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        c3.download_button("🎨 存为 HTML", export_to_pretty_html(curr_chat["messages"], curr_chat["title"], {"system_prompt": curr_chat.get("system_prompt", ""), "model": active_p["model"]}), f"{curr_chat['title']}.html", mime="text/html", use_container_width=True)

        st.divider()
        st.caption("⚠️ *部分手机浏览器会拦截文件下载。若按钮无反应，**请直接长按下方文本框全选复制***：")
        st.text_area("纯文本防拦截备用区", txt_c, height=200, label_visibility="collapsed")

        if st.button("❌ 关闭窗口", use_container_width=True):
            st.rerun()

# ==========================================
# 初始化与水合（修复 1.3：彻底取消 rerun 轮询）
# ==========================================
def _get_default_data():
    default_profiles = [{
        "name": "默认引擎", "base_url": "", "api_key": "", "model": "anthropic/claude-3-5-sonnet-20240620",
        "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 4096,
        "use_top_p": False, "top_p": 1.0, "use_frequency_penalty": False, "frequency_penalty": 0.0
    }]
    first_id = str(uuid.uuid4())
    default_chats = {first_id: {
        "title": "新对话", "messages": [], "session_knowledge": [],
        "system_prompt": "", "is_pinned": False, "is_archived": False
    }}
    return default_profiles, default_chats, first_id

if "initialized" not in st.session_state:
    st.session_state.update({
        "initialized": False,
        "_needs_save": False,
        "is_streaming": False,
        "_render_limit": DEFAULT_RENDER_WINDOW,
        "stop_req": False,
        "_last_save_ts": 0,
        "_last_payload_size": 0,
    })

if not st.session_state.initialized:
    default_profiles, default_chats, first_id = _get_default_data()

    saved_data = None
    if localS is not None:
        try:
            saved_data = localS.getItem(STORAGE_KEY, key="zm_storage")
        except Exception as e:
            print(f"[ZenMux] 读取本地存储失败: {e}")
            saved_data = None

    hydrated = False
    if saved_data not in (None, "", "null"):
        try:
            data = json.loads(saved_data) if isinstance(saved_data, str) else saved_data
            if isinstance(data, dict) and ("profiles" in data or "free_chats" in data):
                st.session_state.profiles = data.get("profiles") or default_profiles
                st.session_state.free_chats = data.get("free_chats") or default_chats
                hydrated = True
        except Exception as e:
            print(f"[ZenMux] 数据解析失败: {e}")

    if not hydrated:
        st.session_state.profiles = default_profiles
        st.session_state.free_chats = default_chats

    if not st.session_state.free_chats:
        st.session_state.free_chats = default_chats

    st.session_state.active_profile_idx = 0
    st.session_state.current_chat_id = list(st.session_state.free_chats.keys())[-1]
    st.session_state.current_page = "💬 自由聊天区"
    st.session_state.initialized = True
    # ⭐ 不 rerun：组件会在下次用户操作时自然同步

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
    if not text:
        return 0
    cn = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
    en = len(re.findall(r'[a-zA-Z]+', text))
    return cn + en

def render_copy_button_inline(text):
    b64 = base64.b64encode(text.encode("utf-8")).decode("utf-8")
    return f'<button class="zm-copy-btn" onclick="window.zmCopy(this, \'{b64}\')">📋 复制</button>'

def extract_file_text(uploaded_file):
    name = uploaded_file.name.lower()
    try:
        if name.endswith('.pdf'):
            return "\n".join([page.extract_text() for page in PdfReader(uploaded_file).pages if page.extract_text()])
        elif name.endswith('.docx'):
            return "\n".join([p.text for p in Document(uploaded_file).paragraphs])
        else:
            return uploaded_file.getvalue().decode('utf-8', errors='ignore')
    except Exception as e:
        return f"文件解析失败: {str(e)}"

def generate_word_doc(messages, is_pure=False):
    doc = Document()
    doc.add_heading('ZenMux 导出文档', 0)
    for msg in messages:
        if msg["role"] == "system": continue
        if is_pure and msg["role"] == "assistant":
            doc.add_paragraph(clean_novel_text(msg["content"]))
        elif not is_pure:
            doc.add_heading("📌 我" if msg["role"]=="user" else "🤖 AI", level=2)
            doc.add_paragraph(msg["content"])
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def export_to_pretty_html(messages, title, meta=None):
    """修复 3.8：移除每条消息的 word_count 调用，大对话导出不再卡顿"""
    if meta is None: meta = {}
    css = "* { margin:0; padding:0; box-sizing:border-box; } body { font-family: -apple-system, sans-serif; background:#f0f2f5; color:#1a1a1a; } .header { background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); color:#fff; padding:24px 32px; position:sticky; top:0; z-index:100; } .header h1 { font-size:22px; font-weight:700; } .header .meta { font-size:12px; opacity:.75; margin-top:6px; } .chat-container { max-width:860px; margin:0 auto; padding:24px 16px 80px; } .info-card { background:#fff; border-radius:12px; padding:20px 24px; margin-bottom:24px; border-left:4px solid #667eea; } .info-row { display:flex; margin-bottom:8px; font-size:13px; line-height:1.6; } .info-label { color:#888; min-width:90px; font-weight:600; } .msg { display:flex; gap:12px; margin-bottom:24px; align-items:flex-start; } .msg.user { flex-direction:row-reverse; } .avatar { width:36px; height:36px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:18px; flex-shrink:0; } .msg.ai .avatar { background:#e8f5e9; } .msg.user .avatar { background:#e3f2fd; } .bubble { max-width:75%; padding:14px 18px; border-radius:16px; line-height:1.8; font-size:15px; white-space:pre-wrap; } .msg.ai .bubble { background:#fff; border-top-left-radius:4px; box-shadow:0 1px 3px rgba(0,0,0,.06);} .msg.user .bubble { background:#d1e7ff; border-top-right-radius:4px; }"
    js = "<script>function toggleInfo() { var el = document.getElementById('infoCard'); el.style.display = (el.style.display==='none') ? 'block' : 'none'; }</script>"

    info_html = ""
    if any(meta.get(k) for k in ["system_prompt", "model"]):
        rows = ""
        if meta.get("model"): rows += f'<div class="info-row"><span class="info-label">🧠 模型</span><span class="info-value">{html.escape(meta["model"])}</span></div>'
        if meta.get("system_prompt"): rows += f'<div class="info-row"><span class="info-label">🎭 人设</span></div><div class="info-value" style="background:#f8f8f8;padding:8px;border-radius:6px;font-size:12px;">{html.escape(meta["system_prompt"])}</div>'
        info_html = f'<div class="info-card" id="infoCard"><h3>⚙️ 配置信息</h3>{rows}</div>'

    # ⭐ 移除 word-count 计算
    msg_html = "".join([f'<div class="msg {"user" if m["role"]=="user" else "ai"}"><div class="avatar">{"🙋‍♂️" if m["role"]=="user" else "🤖"}</div><div><div class="bubble">{html.escape(m["content"]).replace(chr(10), "<br>")}</div></div></div>' for m in messages if m["role"]!="system"])
    date_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    return f"<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>{html.escape(title)}</title><style>{css}</style></head><body><div class='header'><h1>💬 {html.escape(title)}</h1><div class='meta'>{date_str} <span style='text-decoration:underline;cursor:pointer;margin-left:10px' onclick='toggleInfo()'>显示/隐藏配置</span></div></div><div class='chat-container'>{info_html}{msg_html}</div>{js}</body></html>".encode('utf-8')

def fetch_models(base_url, api_key):
    try:
        url = (base_url.strip().rstrip('/') or "https://api.openai.com/v1") + "/models"
        resp = requests.get(url, headers={"Authorization": "Bearer " + api_key.strip()}, timeout=8)
        return (True, sorted([m["id"] for m in resp.json().get("data", [])])) if resp.status_code == 200 else (False, f"状态码 {resp.status_code}")
    except Exception as e:
        return False, str(e)

def get_client():
    p = st.session_state.profiles[st.session_state.active_profile_idx]
    return OpenAI(base_url=p["base_url"].strip() or "https://api.openai.com/v1", api_key=p["api_key"].strip()), p

def build_api_kwargs(profile, api_msgs):
    kw = {"model": profile["model"], "messages": api_msgs, "stream": True}
    if profile.get("use_temperature", True): kw["temperature"] = profile.get("temperature", 0.8)
    if profile.get("use_max_tokens", True):
        kw["max_tokens"] = profile.get("max_tokens", 4096) or 4096
    if profile.get("use_top_p", False): kw["top_p"] = profile.get("top_p", 1.0)
    if profile.get("use_frequency_penalty", False): kw["frequency_penalty"] = profile.get("frequency_penalty", 0.0)
    return kw

def _request_stop():
    st.session_state.stop_req = True

# ==========================================
# 4. 全局侧边栏
# ==========================================
with st.sidebar:
    st.title("🐙 ZenMux")
    page = st.radio("导航", ["💬 自由聊天区", "⚙️ 底层引擎配置"], label_visibility="collapsed")
    st.session_state.current_page = page
    active_p = st.session_state.profiles[st.session_state.active_profile_idx]
    st.caption(f"🟢 当前引擎: {active_p['name']} | 🧠 {active_p['model']}")

    # 修复 3.1 + 4：流式中禁用切换
    is_streaming_now = st.session_state.get("is_streaming", False)
    if is_streaming_now:
        st.warning("🔴 正在生成中，操作已锁定...")

    st.divider()

    if page == "💬 自由聊天区":
        if st.session_state.current_chat_id not in st.session_state.free_chats:
            st.session_state.current_chat_id = list(st.session_state.free_chats.keys())[-1]

        if st.button("➕ 新建对话", use_container_width=True, type="primary", disabled=is_streaming_now):
            nid = str(uuid.uuid4())
            st.session_state.free_chats[nid] = {"title": "新对话", "messages": [], "session_knowledge": [], "system_prompt": "", "is_pinned": False, "is_archived": False}
            st.session_state.current_chat_id = nid
            st.session_state._render_limit = DEFAULT_RENDER_WINDOW
            trigger_save()
            st.rerun()

        with st.expander("📚 历史对话", expanded=True):
            search_q = st.text_input("🔍 搜索", label_visibility="collapsed", placeholder="搜索历史...")
            chat_items = [(cid, cdata) for cid, cdata in st.session_state.free_chats.items() if not cdata.get("is_archived", False) and (not search_q or search_q.lower() in cdata["title"].lower())]
            chat_items.sort(key=lambda x: x[1].get("is_pinned", False), reverse=True)

            for cid, cdata in chat_items:
                prefix = "⭐ " if cid == st.session_state.current_chat_id else ("📌 " if cdata.get("is_pinned") else "📄 ")
                if st.button(prefix + cdata["title"], key=f"sel_{cid}", use_container_width=True, disabled=is_streaming_now):
                    st.session_state.current_chat_id = cid
                    st.session_state._render_limit = DEFAULT_RENDER_WINDOW
                    st.rerun()

            if st.button("🗄️ 归档区", key="toggle_archive", disabled=is_streaming_now):
                st.session_state._show_archive = not st.session_state.get("_show_archive", False)
                st.rerun()
            if st.session_state.get("_show_archive", False):
                for cid, cdata in st.session_state.free_chats.items():
                    if cdata.get("is_archived", False):
                        col1, col2 = st.columns([3, 1])
                        col1.caption(f"📦 {cdata['title']}")
                        if col2.button("恢复", key=f"unarch_{cid}", disabled=is_streaming_now):
                            cdata["is_archived"] = False
                            trigger_save()
                            st.rerun()

        st.divider()
        with st.expander("📦 全量数据快照迁移", expanded=False):
            st.caption("💡 建议每周手动下载一次快照作为灾备")

            # 修复 11：存储用量轻量显示
            _sz = st.session_state.get("_last_payload_size", 0)
            if _sz > 0:
                _sz_mb = _sz / 1024 / 1024
                _pct = min(100, int(_sz / STORAGE_LIMIT_BYTES * 100))
                if _pct >= 80:
                    st.error(f"⚠️ 存储占用 {_sz_mb:.2f} MB ({_pct}%)，请尽快归档！")
                else:
                    st.caption(f"📊 当前占用: {_sz_mb:.2f} MB ({_pct}%)")

            if st.button("🔧 生成快照", use_container_width=True, key="gen_snapshot"):
                st.session_state._snapshot_data = json.dumps(
                    {"profiles": st.session_state.profiles, "free_chats": st.session_state.free_chats},
                    ensure_ascii=False, indent=2
                )

            if st.session_state.get("_snapshot_data"):
                fname = f"ZenMux_Backup_{datetime.now().strftime('%m%d_%H%M')}.json"
                st.download_button("📥 下载快照", st.session_state._snapshot_data.encode('utf-8'),
                                   fname, mime="application/json", use_container_width=True, type="primary")
                with st.expander("📄 防拦截：复制快照代码"):
                    st.code(st.session_state._snapshot_data, language="json")

            if uploaded_ws := st.file_uploader("📂 导入快照 (覆盖当前)", type="json"):
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
    curr_chat = st.session_state.free_chats[st.session_state.current_chat_id]
    current_chat_id = st.session_state.current_chat_id  # 本地变量

    if st.session_state.get("_trigger_export", False):
        if dialog_decorator:
            render_export_modal(curr_chat, active_p)
            st.session_state._trigger_export = False

    # --- 吸顶标题栏 ---
    tc1, tc2 = st.columns([10, 1])
    with tc1:
        # ⭐ 修复 2.1：标题输入框不用 key，避免 widget state 回写覆盖
        new_title = st.text_input("会话标题", curr_chat["title"], label_visibility="collapsed")
        if new_title != curr_chat["title"]:
            st.session_state.free_chats[current_chat_id]["title"] = new_title
            trigger_save()
    with tc2:
        with st.popover("🔽"):
            st.markdown("##### ⚙️ 会话管理")
            btn_c1, btn_c2 = st.columns(2)
            if btn_c1.button("取消置顶" if curr_chat.get("is_pinned") else "📌 置顶", use_container_width=True, key="btn_pin"):
                st.session_state.free_chats[current_chat_id]["is_pinned"] = not curr_chat.get("is_pinned", False)
                trigger_save()
                st.rerun()
            if btn_c2.button("📦 归档", use_container_width=True, key="btn_archive"):
                st.session_state.free_chats[current_chat_id]["is_archived"] = True
                trigger_save()
                st.rerun()
            if btn_c1.button("🗑️ 清空", use_container_width=True, key="btn_clear"):
                st.session_state.free_chats[current_chat_id]["messages"] = []
                st.session_state._render_limit = DEFAULT_RENDER_WINDOW
                trigger_save()
                st.rerun()
            if btn_c2.button("📥 导出", use_container_width=True, key="btn_export"):
                if dialog_decorator:
                    st.session_state._trigger_export = True
                    st.rerun()
                else:
                    st.error("当前 Streamlit 版本过旧，不支持弹窗。")

            st.divider()
            st.markdown("##### 📚 全局设定与知识库")
            # ⭐ 修复 2.1：system_prompt 不用 key
            sp_new = st.text_area("🎭 System Prompt", curr_chat.get("system_prompt", ""), height=80, placeholder="设定此会话专属的全局人设...")
            if sp_new != curr_chat.get("system_prompt", ""):
                st.session_state.free_chats[current_chat_id]["system_prompt"] = sp_new
                trigger_save()

            up_f = st.file_uploader(
                f"📎 常驻参考文件 ({FILE_TYPE_HINT})",
                type=SUPPORTED_FILE_TYPES,
                key=f"kb_{current_chat_id}"
            )
            if up_f:
                content = extract_file_text(up_f)
                if "session_knowledge" not in st.session_state.free_chats[current_chat_id]:
                    st.session_state.free_chats[current_chat_id]["session_knowledge"] = []
                if not any(k["filename"] == up_f.name for k in st.session_state.free_chats[current_chat_id].get("session_knowledge", [])):
                    st.session_state.free_chats[current_chat_id]["session_knowledge"].append({"filename": up_f.name, "content": content})
                    trigger_save()
                    st.rerun()
            for ki, k in enumerate(curr_chat.get("session_knowledge", [])):
                kc1, kc2 = st.columns([4, 1])
                kc1.caption(f"📄 {k['filename']} ({count_words(k['content']):,} 字)")
                if kc2.button("❌", key=f"rm_kb_{current_chat_id}_{ki}"):
                    st.session_state.free_chats[current_chat_id]["session_knowledge"].pop(ki)
                    trigger_save()
                    st.rerun()

    # --- 聊天消息展示 ---
    all_msgs = curr_chat["messages"]
    total_msgs = len(all_msgs)
    render_limit = st.session_state.get("_render_limit", DEFAULT_RENDER_WINDOW)
    start_idx = max(0, total_msgs - render_limit)

    with st.container(border=False):
        editing_idx = st.session_state.get("_editing_chat_idx")

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

            with st.chat_message(msg["role"]):
                if msg.get("files"):
                    for f in msg["files"]:
                        st.caption(f"`{'🔄' if f.get('continuous') else '1️⃣'} 附件: {f['filename']}`")

                if msg["role"] == "user" and editing_idx == i:
                    new_text = st.text_area("✏️ 编辑", msg["content"], key=f"edit_area_{i}", height=100)
                    ebc1, ebc2 = st.columns(2)
                    if ebc1.button("✅ 重新发送", key=f"edit_ok_{i}", type="primary"):
                        st.session_state.free_chats[current_chat_id]["messages"][i]["content"] = new_text
                        st.session_state.free_chats[current_chat_id]["messages"] = st.session_state.free_chats[current_chat_id]["messages"][:i + 1]
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
                        with st.popover("⚙️"):
                            if st.button("✏️ 编辑", key=f"edit_btn_{i}"):
                                st.session_state._editing_chat_idx = i
                                st.rerun()
                            if st.button("🗑️ 删除此条", key=f"del_msg_{i}"):
                                st.session_state.free_chats[current_chat_id]["messages"].pop(i)
                                trigger_save()
                                st.rerun()
                    elif msg["role"] == "assistant":
                        with st.popover("⚙️"):
                            if st.button("🔄 重新生成", key=f"regen_{i}"):
                                st.session_state.free_chats[current_chat_id]["messages"] = st.session_state.free_chats[current_chat_id]["messages"][:i]
                                st.session_state._auto_resend = True
                                trigger_save()
                                st.rerun()
                            if msg.get("_is_half") and st.button("▶️ 续写", key=f"resume_{i}"):
                                st.session_state._resume_idx = i
                                st.session_state._auto_resend = True
                                st.rerun()
                            if st.button("🗑️ 删除此条", key=f"del_msg_ai_{i}"):
                                st.session_state.free_chats[current_chat_id]["messages"].pop(i)
                                trigger_save()
                                st.rerun()
                        wc = count_words(msg["content"])
                        st.markdown(
                            f'<div style="display:flex;justify-content:flex-end;align-items:center;font-size:12px;color:#888;margin-top:4px;">'
                            f'📊 {wc} 字'
                            f'{render_copy_button_inline(msg["content"])}'
                            f'</div>',
                            unsafe_allow_html=True
                        )

    # --- 流式请求处理 ---
    need_resend = st.session_state.pop("_auto_resend", False)
    resume_idx = st.session_state.pop("_resume_idx", None)

    with st.popover("📎 附件"):
        # ⭐ 修复 2.3：dyn_file key 随 chat_id 变化，切换对话自动清空
        dyn_file = st.file_uploader(
            f"跟随消息发送的单次文件 ({FILE_TYPE_HINT})",
            type=SUPPORTED_FILE_TYPES,
            key=f"dyn_file_{current_chat_id}"
        )
        is_continuous = st.checkbox("🔄 持续参考 (勾选后对后续对话一直生效)", value=False, key=f"dyn_file_continuous_{current_chat_id}")

    prompt = st.chat_input("输入消息...")

    if prompt or need_resend:
        if not active_p["api_key"]:
            st.error("⚠️ 请先在【底层引擎配置】中填写 API Key！")
            st.stop()

        # ⭐ 修复 3.1：锁定 chat_id，流式期间即使用户切换也写入正确对话
        locked_chat_id = current_chat_id

        if resume_idx is not None:
            prompt = "请紧接上文最后一个字继续往下写，保持文风和节奏一致。"
            st.session_state.free_chats[locked_chat_id]["messages"][resume_idx]["_is_half"] = False

        if prompt:
            if len(st.session_state.free_chats[locked_chat_id]["messages"]) == 0 and st.session_state.free_chats[locked_chat_id]["title"] == "新对话":
                st.session_state.free_chats[locked_chat_id]["title"] = prompt[:10] + ("..." if len(prompt) > 10 else "")
            new_msg = {"role": "user", "content": prompt}
            if dyn_file and not need_resend:
                new_msg["files"] = [{
                    "filename": dyn_file.name,
                    "content": extract_file_text(dyn_file),
                    "continuous": is_continuous
                }]
            st.session_state.free_chats[locked_chat_id]["messages"].append(new_msg)
            trigger_save()

            with st.chat_message("user"):
                st.markdown(prompt)

        # 构建 API 消息（使用锁定的 chat）
        locked_chat = st.session_state.free_chats[locked_chat_id]
        api_msgs = []
        sp = locked_chat.get("system_prompt", "").strip()
        if sp: api_msgs.append({"role": "system", "content": sp})
        if locked_chat.get("session_knowledge"):
            kb_parts = [f"--- {k['filename']} ---\n{k['content']}" for k in locked_chat["session_knowledge"]]
            api_msgs.append({"role": "system", "content": "【全局参考】：\n" + "\n\n".join(kb_parts)})

        for i, m in enumerate(locked_chat["messages"]):
            content = m["content"]
            if m.get("files"):
                file_texts = [f"--- {f['filename']} ---\n{f['content']}" for f in m["files"] if f.get("continuous") or i == len(locked_chat["messages"]) - 1]
                if file_texts:
                    content = "【附件】\n" + "\n".join(file_texts) + "\n\n【指令】\n" + content
            api_msgs.append({"role": m["role"], "content": content})

        client, profile = get_client()

        with st.chat_message("assistant"):
            # ⭐ 修复 3.3：停止按钮状态清理
            if "stop_gen_btn_fixed" in st.session_state:
                try:
                    del st.session_state["stop_gen_btn_fixed"]
                except Exception:
                    pass
            st.session_state.stop_req = False
            st.session_state.is_streaming = True

            stop_placeholder = st.empty()
            stop_placeholder.button(
                "⏹️ 停止生成",
                key="stop_gen_btn_fixed",
                on_click=_request_stop
            )

            message_placeholder = st.empty()
            full_resp = ""
            try:
                resp = client.chat.completions.create(**build_api_kwargs(profile, api_msgs))
                for chunk in resp:
                    if st.session_state.get("stop_req"):
                        break
                    if chunk.choices and chunk.choices[0].delta.content is not None:
                        full_resp += chunk.choices[0].delta.content
                        message_placeholder.markdown(full_resp + "▌")
                message_placeholder.markdown(full_resp)
                is_half = bool(st.session_state.get("stop_req"))
                if full_resp:
                    # ⭐ 修复 3.1：使用锁定的 chat_id 写入
                    st.session_state.free_chats[locked_chat_id]["messages"].append(
                        {"role": "assistant", "content": full_resp, "_is_half": is_half}
                    )
            except Exception as e:
                st.error(f"请求失败: {str(e)}")
                if full_resp:
                    st.session_state.free_chats[locked_chat_id]["messages"].append(
                        {"role": "assistant", "content": full_resp, "_is_half": True}
                    )
            finally:
                stop_placeholder.empty()
                st.session_state.is_streaming = False
                st.session_state.stop_req = False
                # ⭐ 修复 3.2：同步保存，不依赖 rerun
                st.session_state._needs_save = True
                st.session_state._last_save_ts = 0
                execute_save()
                st.rerun()

# ==========================================
# 模块 2: 底层引擎配置（修复 1.1 / 1.2 / 2.1 / 2.2 核心主诉）
# ==========================================
elif st.session_state.current_page == "⚙️ 底层引擎配置":
    st.header("⚙️ 底层驱动配置")
    p_names = [p["name"] for p in st.session_state.profiles]
    idx = st.radio("切换引擎", range(len(p_names)), format_func=lambda x: p_names[x], index=st.session_state.active_profile_idx, horizontal=True, key="engine_selector")
    st.session_state.active_profile_idx = idx

    if st.button("➕ 新增引擎", use_container_width=True, key="add_engine"):
        st.session_state.profiles.append({
            "name": f"新引擎 {len(p_names) + 1}", "base_url": "", "api_key": "", "model": "anthropic/claude-3-5-sonnet-20240620",
            "use_temperature": True, "temperature": 0.8, "use_max_tokens": True, "max_tokens": 4096,
            "use_top_p": False, "top_p": 1.0, "use_frequency_penalty": False, "frequency_penalty": 0.0
        })
        trigger_save()
        st.rerun()

    st.divider()

    # ⭐ 修复 1.2：始终用完整路径读写，不依赖局部变量 p
    p = st.session_state.profiles[idx]

    c1, c2 = st.columns([3, 1])
    # ⭐ 修复 2.1：text_input 不传 key，避免 widget state 覆盖
    new_name = c1.text_input("引擎标签", p["name"])
    if new_name != p["name"]:
        st.session_state.profiles[idx]["name"] = new_name
        trigger_save()
    if c2.button("💾 保存配置", type="primary", use_container_width=True, key="save_cfg"):
        st.session_state._last_save_ts = 0
        trigger_save()
        st.success("已保存！")

    new_url = st.text_input("Base URL", p["base_url"])
    if new_url != p["base_url"]:
        st.session_state.profiles[idx]["base_url"] = new_url
        trigger_save()

    new_key = st.text_input("API Key", p["api_key"], type="password")
    if new_key != p["api_key"]:
        st.session_state.profiles[idx]["api_key"] = new_key
        trigger_save()

    st.caption("🔒 你的 Key 只保存在浏览器本地缓存中，服务器不会记录。")

    if p["api_key"] and st.button("🔑 测试连通性", key="test_conn"):
        with st.spinner("测试中..."):
            try:
                OpenAI(
                    base_url=p["base_url"].strip() or "https://api.openai.com/v1",
                    api_key=p["api_key"].strip()
                ).chat.completions.create(model=p["model"], messages=[{"role": "user", "content": "Hi"}], max_tokens=5)
                st.success("✅ 连通成功！")
            except Exception as e:
                st.error(f"❌ 失败: {str(e)}")

    m1, m2 = st.columns([3, 1])
    # ⭐ 核心修复 1.1：model 输入框不传 key，彻底解决"刷新模型不生效"
    new_model = m1.text_input("模型映射 (Model ID)", p["model"])
    if new_model != p["model"]:
        st.session_state.profiles[idx]["model"] = new_model
        trigger_save()

    if m2.button("🔄 获取列表", key="fetch_models") and p["api_key"]:
        with st.spinner("获取中..."):
            success, result = fetch_models(p["base_url"], p["api_key"])
            if success and result:
                st.session_state.temp_models = result
                # ⭐ 清理旧的 selectbox 状态，避免残留
                if "model_picker" in st.session_state:
                    try:
                        del st.session_state["model_picker"]
                    except Exception:
                        pass
                st.success(f"✅ 获取到 {len(result)} 个模型！")
            else:
                st.error(f"❌ 失败: {result}")

    # ⭐ 修复 2.1 + 3.10：模型选择后，完整路径写入 + 清理所有相关 widget state
    if "temp_models" in st.session_state:
        sel_m = st.selectbox("选择模型", ["(不覆盖)"] + st.session_state.temp_models, key="model_picker")
        if sel_m != "(不覆盖)":
            st.session_state.profiles[idx]["model"] = sel_m
            # 清理临时数据和 widget state
            del st.session_state.temp_models
            if "model_picker" in st.session_state:
                try:
                    del st.session_state["model_picker"]
                except Exception:
                    pass
            # 强制立即保存
            st.session_state._last_save_ts = 0
            trigger_save()
            execute_save()
            st.rerun()

    with st.expander("🎛️ 运行时超参数", expanded=True):
        # checkbox 保留 key（value 是布尔，无此 bug）
        new_ut = st.checkbox("🔥 Temperature", p.get("use_temperature", True), key=f"ut_{idx}")
        if new_ut != p.get("use_temperature", True):
            st.session_state.profiles[idx]["use_temperature"] = new_ut
            trigger_save()
        if new_ut:
            # ⭐ 修复 2.1：slider 不传 key
            new_t = st.slider("温度", 0.0, 2.0, p.get("temperature", 0.8), 0.1)
            if new_t != p.get("temperature", 0.8):
                st.session_state.profiles[idx]["temperature"] = new_t
                trigger_save()

        new_umt = st.checkbox("📏 Max Tokens", p.get("use_max_tokens", True), key=f"umt_{idx}")
        if new_umt != p.get("use_max_tokens", True):
            st.session_state.profiles[idx]["use_max_tokens"] = new_umt
            trigger_save()
        if new_umt:
            current_mt = p.get("max_tokens", 4096)
            if current_mt > 2000000: current_mt = 2000000
            if current_mt < 1: current_mt = 4096
            # ⭐ 修复 2.1：number_input 不传 key
            new_mt = st.number_input(
                "最大 Token 数",
                min_value=1, max_value=2000000,
                value=current_mt, step=256,
                help=(
                    "支持 1 – 2,000,000。\n\n"
                    "• 常见输出上限：GPT-4o ≈ 16K、Claude 3.5 ≈ 8K、Claude 3.7 ≈ 64K、DeepSeek ≈ 8K\n"
                    "• 大上下文窗口：Claude 200K / GPT-4.1 1M / Gemini 1.5-2.0 Pro 2M"
                )
            )
            if new_mt != current_mt:
                st.session_state.profiles[idx]["max_tokens"] = new_mt
                trigger_save()
            st.caption("快捷预设：")
            preset_cols = st.columns(8)
            presets = [("4K", 4096), ("8K", 8192), ("16K", 16384), ("32K", 32768),
                       ("64K", 65536), ("128K", 131072), ("1M", 1048576), ("2M", 2000000)]
            for _ci, (lbl, val) in enumerate(presets):
                if preset_cols[_ci].button(lbl, key=f"mt_preset_{idx}_{lbl}", use_container_width=True):
                    st.session_state.profiles[idx]["max_tokens"] = val
                    trigger_save()
                    st.rerun()

        new_utp = st.checkbox("🎲 Top P", p.get("use_top_p", False), key=f"utp_{idx}")
        if new_utp != p.get("use_top_p", False):
            st.session_state.profiles[idx]["use_top_p"] = new_utp
            trigger_save()
        if new_utp:
            new_tp = st.slider("Top P", 0.0, 1.0, p.get("top_p", 1.0), 0.05)
            if new_tp != p.get("top_p", 1.0):
                st.session_state.profiles[idx]["top_p"] = new_tp
                trigger_save()

        new_ufp = st.checkbox("🚫 Frequency Penalty", p.get("use_frequency_penalty", False), key=f"ufp_{idx}")
        if new_ufp != p.get("use_frequency_penalty", False):
            st.session_state.profiles[idx]["use_frequency_penalty"] = new_ufp
            trigger_save()
        if new_ufp:
            new_fp = st.slider("惩罚值", -2.0, 2.0, p.get("frequency_penalty", 0.0), 0.1)
            if new_fp != p.get("frequency_penalty", 0.0):
                st.session_state.profiles[idx]["frequency_penalty"] = new_fp
                trigger_save()

    if len(st.session_state.profiles) > 1 and st.button("🗑️ 删除此引擎", type="primary", key="del_engine"):
        st.session_state.profiles.pop(idx)
        st.session_state.active_profile_idx = 0
        trigger_save()
        st.rerun()

# ==========================================
# 统一执行保存
# ==========================================
execute_save()
