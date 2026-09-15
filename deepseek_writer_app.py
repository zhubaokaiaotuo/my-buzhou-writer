import os
import re
import time
from pathlib import Path
from io import BytesIO

import streamlit as st
from docx import Document
from openai import OpenAI

# ============== é¡µé¢é…ç½® ==============
st.set_page_config(
    page_title="å¢¨ç¬”Â·ä¸å‘¨å†™ä½œåŠ©æ‰‹",
    page_icon="âœï¸",
    layout="wide",
)

# ============== å·¥å…·å‡½æ•° ==============

@st.cache_data
def load_docx_text(file_bytes):
    """è¯»å–ä¸Šä¼ çš„ docx æ–‡ä»¶ï¼Œè¿”å›žæ‰€æœ‰æ®µè½"""
    doc = Document(BytesIO(file_bytes))
    return [p.text.strip() for p in doc.paragraphs]


def extract_chapter(paragraphs, chapter_num):
    """æå–æŒ‡å®šç« èŠ‚çš„æ­£æ–‡å’Œç»†çº²"""
    # æ‰¾æ­£æ–‡èµ·ç‚¹
    body_start = 0
    for i, t in enumerate(paragraphs):
        if "ä»¥ä¸Šä¸ºç•ªèŒ„ç‰ˆå¤§çº²" in t or t == "æ­£æ–‡":
            body_start = i
            break

    # æ­£æ–‡ï¼šç²¾ç¡®åŒ¹é… ç¬¬Xç« ï¼Œä¸‹ä¸€ç« åœæ­¢
    body_texts = []
    body_started = False
    chapter_pattern = re.compile(rf"^ç¬¬{chapter_num}ç« ")
    next_pattern = re.compile(r"^ç¬¬\d+ç« ")
    for i in range(body_start, len(paragraphs)):
        t = paragraphs[i]
        if not t:
            continue
        if chapter_pattern.search(t):
            body_started = True
            body_texts.append(t)
            continue
        if body_started:
            if next_pattern.match(t):
                break
            body_texts.append(t)

    # ç»†çº²ï¼šç²¾ç¡®åŒ¹é… ChX æˆ– ç¬¬Xç« ï¼Œä¸‹ä¸€ç« åœæ­¢
    outline_texts = []
    outline_started = False
    outline_pattern = re.compile(rf"^\s*Ch{chapter_num}[^0-9]|^\s*ç¬¬{chapter_num}ç« ")
    next_outline_pattern = re.compile(r"^\s*Ch\d+[^0-9]|^\s*ç¬¬\d+ç« ")
    for i, t in enumerate(paragraphs):
        if outline_pattern.search(t):
            outline_started = True
            outline_texts.append(t)
            continue
        if outline_started:
            if next_outline_pattern.match(t):
                break
            outline_texts.append(t)

    return "\n".join(body_texts), "\n".join(outline_texts)


def count_chinese_words(text):
    return len(re.findall(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]", text))


def build_system_prompt(chapter6_sample, methodology):
    return f"""ä½ æ˜¯ä¸€ä½é•¿ç¯‡ç½‘æ–‡å†™ä½œåŠ©æ‰‹ï¼Œä¸“é—¨æœåŠ¡äºŽç•ªèŒ„æŠ•ç¨¿ç‰ˆå°è¯´ã€Šä¸å‘¨ã€‹çš„åˆ›ä½œã€‚ä½ å«ã€Œå¢¨ç¬”ã€ï¼Œç†Ÿæ‚‰éƒ½å¸‚ç¥žè¯ã€å…‹ç³»æ‚¬ç–‘ã€æ•°æ®è€ƒå¤ç±»åž‹ã€‚

ä½ çš„æ ¸å¿ƒèƒ½åŠ›ï¼š
1. æŒ‰ç»†çº²é€ç« æ‰©å†™æ­£æ–‡
2. æ”¹å†™å·²æœ‰æ­£æ–‡ï¼ŒåŽ»é™¤ AI å‘³
3. æ£€æŸ¥æ­£æ–‡ä¸Žç»†çº²çš„å¯¹åº”å…³ç³»

è¯­è¨€é£Žæ ¼å…³é”®è¯ï¼šç´§å‡‘ã€å¹²ã€å…ˆè§„æ ¼åŽåè¯ï¼›äº”æ„Ÿä¼˜å…ˆäºŽè§£é‡Šï¼›åŠ¨ä½œæ›¿ä»£æƒ…ç»ªæ ‡ç­¾ï¼›è—è§£é‡Šï¼›å¥å­é•¿çŸ­é”™è½ï¼Œæ®µè½å¿½é•¿å¿½çŸ­ã€‚

æ¯ç« å¿…é¡»è‡³å°‘åŒ…å« 1 ä¸ªã€Œæ¯›è¾¹ç»†èŠ‚ã€ï¼šæ— å…³å‰§æƒ…ä½†çœŸå®žçš„ç»†èŠ‚ï¼Œå¦‚éž‹å¸¦ã€å£éŸ³ã€æŒ‡ç”²ç¼é‡Œçš„æ³¥ã€çƒŸå¤´çƒ«ç—•ç­‰ã€‚

åŽ»AIå‘³ç¡¬è§„åˆ™ï¼š
1. çº¢çº¿è¯ï¼ˆå‡ºçŽ°å³åˆ ï¼‰ï¼šå¼‚å¸¸ã€ä»¿ä½›ã€ä¼¼ä¹Žã€å®›å¦‚ã€çŠ¹å¦‚ã€æè‹¥ã€ä¸ç¦ã€ä¸ç”±å¾—ã€å¿ä¸ä½ã€ä¸å¯æ€è®®ã€æžå…¶ã€éžå¸¸ã€æ·¡æ·¡çš„ã€å¾®å¾®çš„ã€å˜´è§’å‹¾èµ·ä¸€æŠ¹å¼§åº¦ã€çœ¼ä¸­é—ªè¿‡ä¸€ä¸ã€é¡¿æ—¶ã€å‡å›ºã€æ­»å¯‚ã€ç‚¸å¼€ã€è¿™ä¸€åˆ»ã€ä¸€æŠ¹ã€ä¸€ä¸ã€æš—è‡ªã€è‹¥æœ‰æ‰€æ€ã€æ¶Œä¸Šå¿ƒå¤´ã€æµ®çŽ°åœ¨è„‘æµ·ã€ç¼“ç¼“ã€è½»è½»ã€‚
2. è­¦æƒ•è¯ï¼ˆæ¯ç« â‰¤2æ¬¡ï¼‰ï¼šçªç„¶ã€çž¬é—´ã€åˆ¹é‚£ã€æŸç§ã€æŸç§è¯´ä¸å‡ºçš„ã€æ·±æ·±åœ°ã€ç´§ç´§åœ°ã€å˜´è§’ã€å¿ƒä¸­ã€æŒ‡èŠ‚ã€çž³å­”ã€çš±äº†çš±çœ‰ã€å¿ƒå¤´ã€çœ¼åº•ã€è„¸ä¸Šæµ®çŽ°ã€çœ¼ç¥žã€æ— è¨€ã€å¾®å¾®ä¸€æ„£ã€æ„£äº†ä¸€ä¸‹ã€ä¸ç”±ã€åªå¾—ã€ä¸ç”±å¾—ã€‚
3. ç¦ç”¨å¤è¯»æ¨¡å¼ï¼šã€Œä¸€ä¸‹ï¼Œä¸€ä¸‹ï¼Œä¸€ä¸‹ã€ç»“å°¾å…¨å·ä¿ç•™ä¸è¶…è¿‡ 1 æ¬¡ï¼›ç¦æ­¢åŒä¸€æ®µè½å†…é‡å¤ä½¿ç”¨åŒä¸€æ¯”å–»ï¼›ç¦æ­¢æ¯æ®µéƒ½æ˜¯ 2-4 å¥çš„å‡åŒ€èŠ‚å¥ã€‚
4. æ›¿æ¢ç¤ºä¾‹ï¼šä¸è¦å†™ã€Œä»–æ„Ÿåˆ°å¾ˆå®³æ€•ã€ï¼Œè¦å†™ã€Œä»–å’½äº†å£å”¾æ²«ï¼Œæ‰‹æŒ‡æŠŠçƒŸç›’ææ‰ã€ï¼›ä¸è¦å†™ã€Œæ¸©åº¦å¼‚å¸¸å‡é«˜ã€ï¼Œè¦å†™ã€Œæ¸©åº¦è®¡æŒ‡é’ˆè·³åˆ°äº†å››åä¸ƒåº¦ã€ã€‚

å‰§æƒ…çºªå¾‹ï¼šä¸¥æ ¼æŒ‰ç»†çº²ä½ç½®é‡Šæ”¾å‰§æƒ…ï¼Œä¸æå‰ï¼›ä¸æ–°å¢žç»†çº²å¤–è§’è‰²ï¼›ä¸æ”¹åŠ¨æ ¸å¿ƒè®¾å®šï¼›æ—¶é—´çº¿ä»¥æœˆåœ†ä¸ºé”šç‚¹ã€‚

è¾“å‡ºæ ¼å¼è¦æ±‚ï¼š
## ç¬¬Xç«  æ ‡é¢˜
[æ­£æ–‡]
---
æœ¬ç« å­—æ•°ï¼šXXXXå­—
çº¢çº¿è¯ï¼š0ä¸ª
è­¦æƒ•è¯ï¼šXä¸ªï¼ˆåˆ—å‡ºï¼‰

ä¸‹é¢æ˜¯ç¬¬6ç« æ”¹å†™èŒƒæœ¬ï¼Œè¯·ä¸¥æ ¼å‚ç…§å…¶è¯­æ„Ÿã€èŠ‚å¥å’Œç»†èŠ‚å¯†åº¦ï¼š

{chapter6_sample}

ä»¥ä¸‹æ˜¯åŽ»AIå‘³æ–¹æ³•è®ºè¡¥å……ï¼š
{methodology}
"""


def call_deepseek(api_key, system_prompt, user_prompt, model="deepseek-chat", temperature=0.7):
    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
        max_tokens=4000,
    )
    return response.choices[0].message.content


def build_user_prompt(chapter_num, mode, body, outline):
    if mode == "æ”¹å†™":
        return f"""è¯·æŒ‰ä»¥ä¸‹è¦æ±‚æ”¹å†™ç¬¬{chapter_num}ç« ï¼š

1. ç»†çº²è¦æ±‚ï¼š
{outline if outline else 'ï¼ˆè¯·ä»ŽçŸ¥è¯†åº“ç»†çº²ä¸­è¯»å–ï¼‰'}

2. å½“å‰æ­£æ–‡ï¼š
{body if body else 'ï¼ˆå½“å‰æ­£æ–‡ä¸ºç©ºï¼Œè¯·æŒ‰ç»†çº²æ‰©å†™ï¼‰'}

3. è¦æ±‚ï¼š
- ä¸¥æ ¼æŒ‰ç»†çº²å½’ä½å‰§æƒ…
- å­—æ•° 2200-2800 å­—
- å½»åº•åŽ» AI å‘³
- å‚ç…§ç¬¬6ç« èŒƒæœ¬é£Žæ ¼
- ç« æœ«é™„å­—æ•°å’Œè¯é¢‘è‡ªæ£€

è¯·ç›´æŽ¥è¾“å‡ºæ”¹å†™åŽçš„å®Œæ•´ç« èŠ‚ã€‚"""
    else:
        return f"""è¯·æŒ‰ä»¥ä¸‹ç»†çº²æ‰©å†™ç¬¬{chapter_num}ç« ï¼š

{outline if outline else 'ï¼ˆè¯·ä»ŽçŸ¥è¯†åº“ç»†çº²ä¸­è¯»å–ï¼‰'}

è¦æ±‚ï¼š
- å­—æ•° 2200-2800 å­—
- å½»åº•åŽ» AI å‘³
- å‚ç…§ç¬¬6ç« èŒƒæœ¬é£Žæ ¼
- ç« æœ«é™„å­—æ•°å’Œè¯é¢‘è‡ªæ£€

è¯·ç›´æŽ¥è¾“å‡ºæ‰©å†™åŽçš„å®Œæ•´ç« èŠ‚ã€‚"""


# ============== çŸ¥è¯†åº“åŠ è½½ ==============
@st.cache_resource
def load_knowledge():
    sample_path = Path(__file__).parent / "ä¸å‘¨_Ch6_æ”¹å†™æ ·ç« .md"
    methodology_path = Path(__file__).parent / "åŽ»AIå‘³é•¿ç¯‡å°è¯´å†™ä½œæ–¹æ³•è®º_report.md"
    chapter6_sample = sample_path.read_text(encoding="utf-8") if sample_path.exists() else "ï¼ˆæœªæ‰¾åˆ°ç¬¬6ç« èŒƒæœ¬ï¼‰"
    methodology = methodology_path.read_text(encoding="utf-8") if methodology_path.exists() else ""
    return build_system_prompt(chapter6_sample, methodology)


# ============== é¡µé¢ UI ==============

st.title("âœï¸ å¢¨ç¬”Â·ä¸å‘¨å†™ä½œåŠ©æ‰‹")
st.caption("åŸºäºŽ DeepSeek API çš„ã€Šä¸å‘¨ã€‹æ”¹å†™/æ‰©å†™å·¥å…·")

# ä¾§è¾¹æ é…ç½®
with st.sidebar:
    st.header("âš™ï¸ é…ç½®")
    api_key = st.text_input("DeepSeek API Key", type="password", help="åœ¨ platform.deepseek.com èŽ·å–")
    model = st.selectbox("æ¨¡åž‹", ["deepseek-chat", "deepseek-reasoner"], index=0)
    temperature = st.slider("Temperature", 0.0, 1.0, 0.7, 0.05)

    st.divider()
    st.markdown("""
    **ä½¿ç”¨æ­¥éª¤ï¼š**
    1. å¡«å…¥ API Key
    2. ä¸Šä¼  docx æºæ–‡ä»¶
    3. é€‰æ‹©ç« èŠ‚å·å’Œæ¨¡å¼
    4. ç‚¹å‡»è¿è¡Œ
    """)

# ä¸»ç•Œé¢
st.header("1. ä¸Šä¼ æºæ–‡ä»¶")
uploaded_file = st.file_uploader("ä¸Šä¼ ã€Šä¸å‘¨ã€‹docx æ–‡ä»¶", type=["docx"])

if uploaded_file is not None:
    paragraphs = load_docx_text(uploaded_file.getvalue())
    st.success(f"æ–‡ä»¶å·²è¯»å–ï¼Œå…± {len(paragraphs)} æ®µ")

    # è‡ªåŠ¨æå–æ‰€æœ‰ç« èŠ‚å·ï¼Œç”¨äºŽæ ¡éªŒ
    chapter_numbers = sorted(set(
        int(m.group(1))
        for t in paragraphs
        for m in [re.search(r"ç¬¬(\d+)ç« ", t)]
        if m
    ))
    if chapter_numbers:
        st.info(f"æ£€æµ‹åˆ°ç« èŠ‚ï¼šç¬¬ {chapter_numbers[0]} ç«  è‡³ ç¬¬ {chapter_numbers[-1]} ç« ")

    system_prompt = load_knowledge()

    # ============== å•ç« æ”¹å†™ ==============
    st.header("2. å•ç« æ”¹å†™")
    col1, col2 = st.columns(2)
    with col1:
        single_chapter = st.number_input("ç« èŠ‚å·", min_value=1, max_value=200, value=7, step=1, key="single_ch")
    with col2:
        single_mode = st.selectbox("æ¨¡å¼", ["æ”¹å†™", "æ‰©å†™"], index=0, key="single_mode")

    if st.button("ðŸš€ å¼€å§‹ç”Ÿæˆ", type="primary", key="single_run"):
        if not api_key:
            st.error("è¯·å…ˆå¡«å†™ DeepSeek API Key")
        else:
            with st.spinner("æ­£åœ¨æå–ç« èŠ‚å†…å®¹..."):
                body, outline = extract_chapter(paragraphs, int(single_chapter))

            if not outline:
                st.warning("æœªæ‰¾åˆ°è¯¥ç« èŠ‚ç»†çº²ï¼Œå°†ç›´æŽ¥æŒ‰åŽŸæ–‡æ”¹å†™")
            if not body:
                st.warning("æœªæ‰¾åˆ°è¯¥ç« èŠ‚æ­£æ–‡ï¼Œå°†æŒ‰ç»†çº²æ‰©å†™")

            user_prompt = build_user_prompt(int(single_chapter), single_mode, body, outline)

            with st.spinner("æ­£åœ¨è°ƒç”¨ DeepSeekï¼Œå¤§æ¦‚éœ€è¦ 30-60 ç§’..."):
                try:
                    result = call_deepseek(api_key, system_prompt, user_prompt, model, temperature)
                    word_count = count_chinese_words(result)

                    st.header("3. ç”Ÿæˆç»“æžœ")
                    st.text_area("æ­£æ–‡", result, height=500)
                    st.info(f"ä¸­æ–‡å­—æ•°çº¦ï¼š{word_count}")

                    st.download_button(
                        label="ðŸ“¥ ä¸‹è½½ä¸º .md",
                        data=result.encode("utf-8"),
                        file_name=f"ä¸å‘¨_Ch{single_chapter}_{single_mode}.md",
                        mime="text/markdown",
                    )
                except Exception as e:
                    st.error(f"ç”Ÿæˆå¤±è´¥ï¼š{e}")

    st.divider()

    # ============== æ‰¹é‡æ”¹å†™ ==============
    st.header("3. æ‰¹é‡æ”¹å†™ / å¯¼å‡º")
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        batch_start = st.number_input("èµ·å§‹ç« èŠ‚", min_value=1, max_value=200, value=7, step=1, key="batch_start")
    with col_b2:
        batch_end = st.number_input("ç»“æŸç« èŠ‚", min_value=1, max_value=200, value=17, step=1, key="batch_end")
    with col_b3:
        batch_size = st.number_input("æ¯æ‰¹æ•°é‡", min_value=1, max_value=5, value=3, step=1,
                                     help="Streamlit Cloud å•æ¬¡è¿è¡Œå¤ªé•¿ä¼šè¶…æ—¶ï¼Œå»ºè®®æ¯æ‰¹ 3 ç« ")

    batch_mode = st.selectbox("æ‰¹é‡æ¨¡å¼", ["æ”¹å†™", "æ‰©å†™"], index=0, key="batch_mode")

    if st.button("ðŸš€ æ‰¹é‡æ”¹å†™", type="primary", key="batch_run"):
        if not api_key:
            st.error("è¯·å…ˆå¡«å†™ DeepSeek API Key")
        elif batch_end < batch_start:
            st.error("ç»“æŸç« èŠ‚ä¸èƒ½å°äºŽèµ·å§‹ç« èŠ‚")
        else:
            results = {}
            progress_bar = st.progress(0)
            status = st.empty()

            chapters_to_run = list(range(batch_start, batch_end + 1))
            total = len(chapters_to_run)

            for idx, ch in enumerate(chapters_to_run):
                status.text(f"æ­£åœ¨å¤„ç†ç¬¬ {ch} ç« ... ({idx + 1}/{total})")
                body, outline = extract_chapter(paragraphs, ch)
                user_prompt = build_user_prompt(ch, batch_mode, body, outline)

                try:
                    result = call_deepseek(api_key, system_prompt, user_prompt, model, temperature)
                    results[ch] = result
                except Exception as e:
                    results[ch] = f"ã€ç”Ÿæˆå¤±è´¥ã€‘{e}"
                    st.error(f"ç¬¬ {ch} ç« å¤±è´¥ï¼š{e}")

                progress_bar.progress((idx + 1) / total)

            st.session_state.batch_results = results
            st.success(f"æ‰¹é‡æ”¹å†™å®Œæˆï¼Œå…± {len(results)} ç« ")

    # æ˜¾ç¤ºæ‰¹é‡ç»“æžœå’Œä¸‹è½½
    if "batch_results" in st.session_state and st.session_state.batch_results:
        results = st.session_state.batch_results

        with st.expander("æŸ¥çœ‹æ‰¹é‡ç»“æžœ"):
            for ch in sorted(results.keys()):
                st.subheader(f"ç¬¬ {ch} ç« ")
                st.text_area(f"ç¬¬{ch}ç« ç»“æžœ", results[ch], height=300, key=f"result_ch_{ch}")

        # åˆå¹¶ä¸‹è½½
        combined_md = "\n\n---\n\n".join(
            [f"{results[ch]}" for ch in sorted(results.keys())]
        )

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                label="ðŸ“¥ ä¸‹è½½å…¨éƒ¨ç»“æžœ.md",
                data=combined_md.encode("utf-8"),
                file_name=f"ä¸å‘¨_Ch{batch_start}-{batch_end}_{batch_mode}.md",
                mime="text/markdown",
            )
        with col_d2:
            # ç”Ÿæˆ docx
            doc = Document()
            for ch in sorted(results.keys()):
                doc.add_paragraph(results[ch])
                doc.add_paragraph()
            docx_io = BytesIO()
            doc.save(docx_io)
            docx_io.seek(0)
            st.download_button(
                label="ðŸ“¥ ä¸‹è½½å…¨éƒ¨ç»“æžœ.docx",
                data=docx_io.getvalue(),
                file_name=f"ä¸å‘¨_Ch{batch_start}-{batch_end}_{batch_mode}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

else:
    st.info("è¯·ä¸Šä¼  docx æ–‡ä»¶åŽå¼€å§‹")
    
