import streamlit as st
from openai import OpenAI
import io, json, re, requests, uuid, html, time
from datetime import datetime
from docx import Document
from pypdf import PdfReader
from streamlit_local_storage import LocalStorage

# ==========================================
# 1. 页面全局配置与精准 UI 优化
# ==========================================
st.set_page_config(page_title="ZenMux 创作者工作站", page_icon="🐙", layout="wide")
st.markdown("""
    <style>
    /* 1. 精准隐藏右上角 Github 标志和部署按钮，绝不误伤左侧菜单键 */
    [data-testid="stToolbar"], .stAppDeployButton { display: none !important; }
    
    /* 2. 优化按钮圆角 */
    .stButton>button { border-radius: 8px; transition: all 0.2s; }
    
    /* 3. 手机端与电脑端安全边距适配，保证汉堡菜单永远可见 */
    .block-container { padding-top: 2rem !important; padding-bottom: 5rem !important; }
    @media (max-width: 768px) {
        .block-container { padding-top: 3.5rem !important; } /* 给手机顶部留出菜单键空间 */
    }
    
    /* 4. 优化输入框上方的附件小加号 */
    [data-testid="stPopover"] { margin-bottom: -15px; z-index: 10; }
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
        data = {"profiles": st.session_state.profiles, "free_chats": st.session_state.free_chats}
        if localS:
            try: localS.setItem("zenmux_data", json.dumps(data))
            except Exception: pass
        st.session_state._needs_save = False

if "initialized" not in st.session_state:
    st.session_state.update({"initialized": False, "ls_loaded": False, "_needs_save": False, "is_streaming": False, "ls_wait_count": 0})

if not st.session_state.ls_loaded:
    saved_data = None
    if localS:
        try: saved_data = localS.getItem("zenmux_data")
        except Exception: pass
            
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

def count_words(text): return len(re.findall(r'[\u4e00-\u9fff]', text)) + len(re.findall(r'[a-zA-Z]+', text))

def extract_file_text(uploaded_file):
    name = uploaded_file.name.lower()
    try:
        if name.endswith('.pdf'): return "\n".join([page.extract_text() for page in PdfReader(uploaded_file).pages if page.extract_text()])
        elif name.endswith('.docx'): return "\n".join([p.text for p in Document(uploaded_file).paragraphs])
        else: return uploaded_file.getvalue().decode('utf-8', errors='ignore')
    except Exception as e: return f"文件解析失败: {str(e)}"

def generate_word_doc(messages, is_pure=False):
    doc = Document()
    doc.add_heading('ZenMux 导出文档', 0)
    for msg in messages:
        if msg["role"] == "system": continue
        if is_pure and msg["role"] == "assistant": doc.add_paragraph(clean_novel_text(msg["content"]))
        elif not is_pure:
            doc.add_heading("📌 我" if msg["role"]=="user" else "🤖 AI", level=2)
            doc.add_paragraph(msg["content"])
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def export_to_pretty_html(messages, title, meta=None):
    if meta is None: meta = {}
    css = "* { margin:0; padding:0; box-sizing:border-box; } body { font-family: -apple-system, sans-serif; background:#f0f2f5; color:#1a1a1a; } .header { background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); color:#fff; padding:24px 32px; position:sticky; top:0; z-index:100; } .header h1 { font-size:22px; font-weight:700; } .header .meta { font-size:12px; opacity:.75; margin-top:6px; } .chat-container { max-width:860px; margin:0 auto; padding:24px 16px 80px; } .info-card { background:#fff; border-radius:12px; padding:20px 24px; margin-bottom:24px; border-left:4px solid #667eea; } .info-row { display:flex; margin-bottom:8px; font-size:13px; line-height:1.6; } .info-label { color:#888; min-width:90px; font-weight:600; } .msg { display:flex; gap:12px; margin-bottom:24px; align-items:flex-start; } .msg.user { flex-direction:row-reverse; } .avatar { width:36px; height:36px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:18px; flex-shrink:0; } .msg.ai .avatar { background:#e8f5e9; } .msg.user .avatar { background:#e3f2fd; } .bubble { max-width:75%; padding:14px 18px; border-radius:16px; line-height:1.8; font-size:15px; white-space:pre-wrap; } .msg.ai .bubble { background:#fff; border-top-left-radius:4px; box-shadow:0 1px 3px rgba(0,0,0,.06);} .msg.user .bubble { background:#d1e7ff; border-top-right-radius:4px; } .word-count { font-size:11px; color:#999; margin-top:6px; text-align:right; } .msg.user .word-count { text-align:left; }"
    js = "<script>function toggleInfo() { var el = document.getElementById('infoCard'); el.style.display = (el.style.display==='none') ? 'block' : 'none'; }</script>"
    
    info_html = ""
    if any(meta.get(k) for k in ["system_prompt", "model"]):
        rows = ""
        if meta.get("model"): rows += f'<div class="info-row"><span class="info-label">🧠 模型</span><span class="info-value">{meta["model"]}</span></div>'
        if meta.get("system_prompt"): rows += f'<div class="info-row"><span class="info-label">🎭 人设</span></div><div class="info-value" style="background:#f8f8f8;padding:8px;border-radius:6px;font-size:12px;">{html.escape(meta["system_prompt"])}</div>'
        info_html = f'<div class="info-card" id="infoCard"><h3>⚙️ 配置信息</h3>{rows}</div>'

    msg_html = "".join([f'<div class="msg {"user" if m["role"]=="user" else "ai"}"><div class="avatar">{"🙋‍♂️" if m["role"]=="user" else "🤖"}</div><div><div class="bubble">{html.escape(m["content"]).replace(chr(10), "<br>")}</div><div class="word-count">{count_words(m["content"])} 字</div></div></div>' for m in messages if m["role"]!="system"])
    date_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    return f"<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>{title}</title><style>{css}</style></head><body><div class='header'><h1>💬 {title}</h1><div class='meta'>{date_str} <span style='text-decoration:underline;cursor:pointer;margin-left:10px' onclick='toggleInfo()'>显示/隐藏配置</span></div></div><div class='chat-container'>{info_html}{msg_html}</div>{js}</body></html>".encode('utf-8')

def fetch_models(base_url, api_key):
    try:
        url = (base_url.strip().rstrip('/') or "https://api.openai.com/v1") + "/models"
        resp = requests.get(url, headers={"Authorization": "Bearer " + api_key.strip()}, timeout=8)
        return (True, sorted([m["id"] for m in resp.json().get("data", [])])) if resp.status_code == 200 else (False, f"状态码 {resp.status_code}")
    except Exception as e: return False, str(e)

def get_client():
    p = st.session_state.profiles[st.session_state.active_profile_idx]
    return OpenAI(base_url=p["base_url"].strip() or "https://api.openai.com/v1", api_key=p["api_key"].strip()), p

def build_api_kwargs(profile, api_msgs):
    kw = {"model": profile["model"], "messages": api_msgs, "stream": True}
    for key in ["temperature", "max_tokens", "top_p", "frequency_penalty"]:
        if profile.get(f"use_{key}", key in ["temperature", "max_tokens"]): kw[key] = profile.get(key)
    return kw

# ==========================================
# 4. 全局侧边栏 (控制台与历史管理)
# ==========================================
with st.sidebar:
    st.title("🐙 ZenMux")
    page = st.radio("导航", ["💬 自由聊天区", "⚙️ 底层引擎配置"], label_visibility="collapsed")
    st.session_state.current_page = page
    active_p = st.session_state.profiles[st.session_state.active_profile_idx]
    st.caption(f"🟢 当前引擎: {active_p['name']} | 🧠 {active_p['model']}")
    st.divider()

    if page == "💬 自由聊天区":
        # 当前会话指针防护
        if st.session_state.current_chat_id not in st.session_state.free_chats:
            st.session_state.current_chat_id = list(st.session_state.free_chats.keys())[-1]
        curr_chat = st.session_state.free_chats[st.session_state.current_chat_id]

        if st.button("➕ 新建对话", use_container_width=True, type="primary"):
            nid = str(uuid.uuid4())
            st.session_state.free_chats[nid] = {"title": "新对话", "messages": [], "session_knowledge": [], "system_prompt": "", "is_pinned": False, "is_archived": False}
            st.session_state.current_chat_id = nid
            trigger_save()
            st.rerun()

        # 分组1：历史对话
        with st.expander("📚 历史对话", expanded=True):
            search_q = st.text_input("🔍 搜索", label_visibility="collapsed", placeholder="搜索历史...")
            chat_items = [(cid, cdata) for cid, cdata in st.session_state.free_chats.items() if not cdata.get("is_archived", False) and (not search_q or search_q.lower() in cdata["title"].lower())]
            chat_items.sort(key=lambda x: x[1].get("is_pinned", False), reverse=True)

            for cid, cdata in chat_items:
                prefix = "⭐ " if cid == st.session_state.current_chat_id else ("📌 " if cdata.get("is_pinned") else "📄 ")
                if st.button(prefix + cdata["title"], key=f"sel_{cid}", use_container_width=True):
                    st.session_state.current_chat_id = cid
                    st.rerun()
            
            if st.button("🗄️ 归档区"):
                st.session_state._show_archive = not st.session_state.get("_show_archive", False)
                st.rerun()
            if st.session_state.get("_show_archive", False):
                for cid, cdata in st.session_state.free_chats.items():
                    if cdata.get("is_archived", False):
                        col1, col2 = st.columns([3, 1])
                        col1.caption(f"📦 {cdata['title']}")
                        if col2.button("恢复", key=f"unarch_{cid}"):
                            cdata["is_archived"] = False
                            trigger_save()
                            st.rerun()

        # 分组2：当前对话全局设定 (万物归一)
        with st.expander("⚙️ 当前对话设定", expanded=False):
            new_title = st.text_input("✏️ 重命名此对话", curr_chat["title"])
            if new_title != curr_chat["title"]:
                curr_chat["title"] = new_title
                trigger_save()
                
            curr_chat["system_prompt"] = st.text_area("🎭 全局人设 (System Prompt)", curr_chat.get("system_prompt", ""), height=80)
            
            up_f = st.file_uploader("📎 添加常驻知识库", type=['txt', 'md', 'pdf', 'docx'], key=f"kb_{st.session_state.current_chat_id}")
            if up_f:
                if not any(k["filename"] == up_f.name for k in curr_chat.get("session_knowledge", [])):
                    if "session_knowledge" not in curr_chat: curr_chat["session_knowledge"] = []
                    curr_chat["session_knowledge"].append({"filename": up_f.name, "content": extract_file_text(up_f)})
                    trigger_save()
                    st.rerun()
            for ki, k in enumerate(curr_chat.get("session_knowledge", [])):
                k1, k2 = st.columns([4, 1])
                k1.caption(f"📄 {k['filename']}")
                if k2.button("❌", key=f"rm_kb_{ki}"):
                    curr_chat["session_knowledge"].pop(ki)
                    trigger_save()
                    st.rerun()
                    
            st.divider()
            if st.button("取消置顶" if curr_chat.get("is_pinned") else "📌 置顶此会话", use_container_width=True):
                curr_chat["is_pinned"] = not curr_chat.get("is_pinned", False)
                trigger_save()
                st.rerun()
            if st.button("📦 归档此会话", use_container_width=True):
                curr_chat["is_archived"] = True
                trigger_save()
                st.rerun()
            if st.button("🗑️ 清空聊天记录", use_container_width=True):
                curr_chat["messages"] = []
                trigger_save()
                st.rerun()
            if st.button("📥 导出聊天记录", use_container_width=True):
                st.session_state._show_export = not st.session_state.get("_show_export", False)
                st.rerun()

        # 数据快照备份
        st.divider()
        with st.expander("📦 全量数据快照迁移", expanded=False):
            full_data = json.dumps({"profiles": st.session_state.profiles, "free_chats": st.session_state.free_chats}, ensure_ascii=False, indent=2).encode('utf-8')
            st.download_button("📥 导出全量快照", full_data, f"ZenMux_Backup_{datetime.now().strftime('%m%d_%H%M')}.json", "application/json", use_container_width=True, type="primary")
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
                except: st.error("导入失败")

# ==========================================
# 模块 1: 自由聊天区 (纯净右侧主战场)
# ==========================================
if st.session_state.current_page == "💬 自由聊天区":
    curr_chat = st.session_state.free_chats[st.session_state.current_chat_id]
    
    # 顶部仅保留固定的标题
    st.markdown(f"<h3 style='margin-top:-10px; padding-bottom:15px; border-bottom:1px solid #ddd;'>💬 {curr_chat['title']}</h3>", unsafe_allow_html=True)

    # 导出面板（如果侧边栏点击了导出）
    if st.session_state.get("_show_export", False):
        with st.container(border=True):
            exp_mode = st.radio("导出格式", ["完整记录", "纯享正文"], horizontal=True, label_visibility="collapsed")
            is_pure = (exp_mode == "纯享正文")
            txt_c = "\n\n".join([clean_novel_text(m['content']) for m in curr_chat["messages"] if m['role'] == 'assistant']) if is_pure else "\n".join([f"{'我' if m['role']=='user' else 'AI'}:\n{m['content']}\n\n{'-'*40}\n" for m in curr_chat["messages"]])
            ec1, ec2, ec3 = st.columns(3)
            ec1.download_button("📥 TXT", txt_c.encode('utf-8'), f"{curr_chat['title']}.txt", use_container_width=True)
            ec2.download_button("📥 Word", generate_word_doc(curr_chat["messages"], is_pure), f"{curr_chat['title']}.docx", use_container_width=True)
            ec3.download_button("🎨 HTML", export_to_pretty_html(curr_chat["messages"], curr_chat["title"], {"system_prompt": curr_chat.get("system_prompt", ""), "model": active_p["model"]}), f"{curr_chat['title']}.html", "text/html", use_container_width=True)

    # 无限向下滚动的聊天区域
    with st.container(border=False):
        editing_idx = st.session_state.get("_editing_chat_idx")
        for i, msg in enumerate(curr_chat["messages"]):
            if msg["role"] == "system": continue
            
            with st.chat_message(msg["role"]):
                if msg.get("files"):
                    for f in msg["files"]: st.caption(f"`{'🔄' if f.get('continuous') else '1️⃣'} 附件: {f['filename']}`")
                        
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
                        with st.popover("⚙️"):
                            if st.button("✏️ 编辑", key=f"edit_btn_{i}"):
                                st.session_state._editing_chat_idx = i
                                st.rerun()
                            if st.button("🗑️ 删除此条", key=f"del_msg_{i}"):
                                curr_chat["messages"].pop(i)
                                trigger_save()
                                st.rerun()
                    elif msg["role"] == "assistant":
                        with st.popover("⚙️"):
                            if st.button("🔄 重新生成", key=f"regen_{i}"):
                                curr_chat["messages"] = curr_chat["messages"][:i]
                                st.session_state._auto_resend = True
                                trigger_save()
                                st.rerun()
                            if msg.get("_is_half") and st.button("▶️ 续写", key=f"resume_{i}"):
                                st.session_state._resume_idx = i
                                st.session_state._auto_resend = True
                                st.rerun()
                            if st.button("🗑️ 删除此条", key=f"del_msg_ai_{i}"):
                                curr_chat["messages"].pop(i)
                                trigger_save()
                                st.rerun()
                        st.caption(f"📊 {count_words(msg['content'])} 字")

    need_resend = st.session_state.pop("_auto_resend", False)
    resume_idx = st.session_state.pop("_resume_idx", None)
    
    # 底部输入框与附件按钮
    with st.popover("➕"):
        dyn_file = st.file_uploader("跟随消息发送附件", type=['txt', 'md', 'pdf', 'docx'], key="dyn_file")
        is_continuous = st.checkbox("🔄 持续参考 (勾选后后续对话一直生效)", value=False)

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
            with st.chat_message("user"): st.markdown(prompt)

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
                for chunk in client.chat.completions.create(**build_api_kwargs(profile, api_msgs)):
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
                OpenAI(base_url=p["base_url"].strip() or "https://api.openai.com/v1", api_key=p["api_key"].strip()).chat.completions.create(model=p["model"], messages=[{"role": "user", "content": "Hi"}], max_tokens=5)
                st.success("✅ 连通成功！")
            except Exception as e: st.error(f"❌ 失败: {str(e)}")

    m1, m2 = st.columns([3, 1])
    p["model"] = m1.text_input("模型映射 (Model ID)", p["model"])
    if m2.button("🔄 获取列表") and p["api_key"]:
        with st.spinner("获取中..."):
            success, result = fetch_models(p["base_url"], p["api_key"])
            if success and result:
                st.session_state.temp_models = result
                st.success(f"✅ 获取到 {len(result)} 个模型！")
            else: st.error(f"❌ 失败: {result}")

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
