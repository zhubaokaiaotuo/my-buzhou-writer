import os
import re
import time
from pathlib import Path

import streamlit as st
from docx import Document
from openai import OpenAI

# ============== 页面配置 ==============
st.set_page_config(
    page_title="墨笔·不周写作助手",
    page_icon="✍️",
    layout="wide",
)

# ============== 工具函数 ==============

@st.cache_data
def load_docx_text(file_bytes):
    """读取上传的 docx 文件，返回所有段落"""
    from io import BytesIO
    doc = Document(BytesIO(file_bytes))
    return [p.text.strip() for p in doc.paragraphs]


def extract_chapter(paragraphs, chapter_num):
    """提取指定章节的正文和细纲"""
    # 找正文起点
    body_start = 0
    for i, t in enumerate(paragraphs):
        if "以上为番茄版大纲" in t or t == "正文":
            body_start = i
            break

    # 正文
    body_texts = []
    body_started = False
    chapter_pattern = re.compile(rf"^第{chapter_num}章|^第[一二三四五六七八九十百]+章")
    for i in range(body_start, len(paragraphs)):
        t = paragraphs[i]
        if not t:
            continue
        if chapter_pattern.search(t):
            body_started = True
            body_texts.append(t)
            continue
        if body_started:
            if re.match(r"^第\d+章|^第[一二三四五六七八九十百]+章", t):
                break
            body_texts.append(t)

    # 细纲
    outline_texts = []
    outline_started = False
    outline_pattern = re.compile(rf"^\s*Ch{chapter_num}[^0-9]|^\s*第{chapter_num}章")
    for i, t in enumerate(paragraphs):
        if outline_pattern.search(t):
            outline_started = True
            outline_texts.append(t)
            continue
        if outline_started:
            if re.match(r"^\s*Ch\d+[^0-9]|^\s*第\d+章", t):
                break
            outline_texts.append(t)

    return "\n".join(body_texts), "\n".join(outline_texts)


def count_chinese_words(text):
    return len(re.findall(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]", text))


def build_system_prompt(chapter6_sample, methodology, guide):
    return f"""你是一位长篇网文写作助手，专门服务于番茄投稿版小说《不周》的创作。你叫「墨笔」，熟悉都市神话、克系悬疑、数据考古类型。

你的核心能力：
1. 按细纲逐章扩写正文
2. 改写已有正文，去除 AI 味
3. 检查正文与细纲的对应关系

语言风格关键词：紧凑、干、先规格后名词；五感优先于解释；动作替代情绪标签；藏解释；句子长短错落，段落忽长忽短。

每章必须至少包含 1 个「毛边细节」：无关剧情但真实的细节，如鞋带、口音、指甲缝里的泥、烟头烫痕等。

去AI味硬规则：
1. 红线词（出现即删）：异常、仿佛、似乎、宛如、犹如、恍若、不禁、不由得、忍不住、不可思议、极其、非常、淡淡的、微微的、嘴角勾起一抹弧度、眼中闪过一丝、顿时、凝固、死寂、炸开、这一刻、一抹、一丝、暗自、若有所思、涌上心头、浮现在脑海、缓缓、轻轻。
2. 警惕词（每章≤2次）：突然、瞬间、刹那、某种、某种说不出的、深深地、紧紧地、嘴角、心中、指节、瞳孔、皱了皱眉、心头、眼底、脸上浮现、眼神、无言、微微一愣、愣了一下、不由、只得、不由得。
3. 禁用复读模式：「一下，一下，一下」结尾全卷保留不超过 1 次；禁止同一段落内重复使用同一比喻；禁止每段都是 2-4 句的均匀节奏。
4. 替换示例：不要写「他感到很害怕」，要写「他咽了口唾沫，手指把烟盒捏扁」；不要写「温度异常升高」，要写「温度计指针跳到了四十七度」。

剧情纪律：严格按细纲位置释放剧情，不提前；不新增细纲外角色；不改动核心设定；时间线以月圆为锚点。

输出格式要求：
## 第X章 标题
[正文]
---
本章字数：XXXX字
红线词：0个
警惕词：X个（列出）

下面是第6章改写范本，请严格参照其语感、节奏和细节密度：

{chapter6_sample}

以下是去AI味方法论补充：
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


# ============== 页面 UI ==============

st.title("✍️ 墨笔·不周写作助手")
st.caption("基于 DeepSeek API 的《不周》改写/扩写工具")

# 侧边栏配置
with st.sidebar:
    st.header("⚙️ 配置")
    api_key = st.text_input("DeepSeek API Key", type="password", help="在 platform.deepseek.com 获取")
    model = st.selectbox("模型", ["deepseek-chat", "deepseek-reasoner"], index=0)
    temperature = st.slider("Temperature", 0.0, 1.0, 0.7, 0.05)

    st.divider()
    st.markdown("""
    **使用步骤：**
    1. 填入 API Key
    2. 上传 docx 源文件
    3. 选择章节号和模式
    4. 点击运行
    """)

# 主界面
st.header("1. 上传源文件")
uploaded_file = st.file_uploader("上传《不周》docx 文件", type=["docx"])

if uploaded_file is not None:
    paragraphs = load_docx_text(uploaded_file.getvalue())
    st.success(f"文件已读取，共 {len(paragraphs)} 段")

    st.header("2. 选择章节")
    col1, col2 = st.columns(2)
    with col1:
        chapter_num = st.number_input("章节号", min_value=1, max_value=60, value=7, step=1)
    with col2:
        mode = st.selectbox("模式", ["改写", "扩写"], index=0)

    if st.button("🚀 开始生成", type="primary"):
        if not api_key:
            st.error("请先填写 DeepSeek API Key")
        else:
            with st.spinner("正在提取章节内容..."):
                body, outline = extract_chapter(paragraphs, int(chapter_num))

            if not outline:
                st.warning("未找到该章节细纲，将直接按原文改写")

            # 加载知识库
            sample_path = Path(__file__).parent / "不周_Ch6_改写样章.md"
            methodology_path = Path(__file__).parent / "去AI味长篇小说写作方法论_report.md"

            chapter6_sample = sample_path.read_text(encoding="utf-8") if sample_path.exists() else "（未找到第6章范本）"
            methodology = methodology_path.read_text(encoding="utf-8") if methodology_path.exists() else ""

            system_prompt = build_system_prompt(chapter6_sample, methodology, "")

            if mode == "改写":
                user_prompt = f"""请按以下要求改写第{chapter_num}章：

1. 细纲要求：
{outline if outline else '（请从知识库细纲中读取）'}

2. 当前正文：
{body if body else '（当前正文为空，请按细纲扩写）'}

3. 要求：
- 严格按细纲归位剧情
- 字数 2200-2800 字
- 彻底去 AI 味
- 参照第6章范本风格
- 章末附字数和词频自检

请直接输出改写后的完整章节。"""
            else:
                user_prompt = f"""请按以下细纲扩写第{chapter_num}章：

{outline if outline else '（请从知识库细纲中读取）'}

要求：
- 字数 2200-2800 字
- 彻底去 AI 味
- 参照第6章范本风格
- 章末附字数和词频自检

请直接输出扩写后的完整章节。"""

            with st.spinner("正在调用 DeepSeek，大概需要 30-60 秒..."):
                try:
                    result = call_deepseek(api_key, system_prompt, user_prompt, model, temperature)
                    word_count = count_chinese_words(result)

                    st.header("3. 生成结果")
                    st.text_area("正文", result, height=500)
                    st.info(f"中文字数约：{word_count}")

                    # 下载按钮
                    st.download_button(
                        label="📥 下载为 .md",
                        data=result.encode("utf-8"),
                        file_name=f"不周_Ch{chapter_num}_{mode}.md",
                        mime="text/markdown",
                    )
                except Exception as e:
                    st.error(f"生成失败：{e}")

else:
    st.info("请上传 docx 文件后开始")
