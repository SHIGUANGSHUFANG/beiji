"""
app.py  ──  背记系统 Streamlit 主程序
=============================================
三个学习步骤：
  Step 1  知识展示    直接看知识原文
  Step 2  气泡池填空  拖拽正确词到下划线空格
  Step 3  快速练习    逐题单选 + 即时反馈

上传 Word(.docx) 后自动解析；未上传时使用演示数据。
"""

import random
import re
import streamlit as st

from parser import parse_docx, generate_distractors, DEMO_ITEMS, DEEPSEEK_API_KEY

# ──────────────────────────────────────────────
# 页面基础配置
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="背记系统",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# 全局 CSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
/* 整体字体与背景 */
html, body, [class*="css"] {
    font-family: "Noto Sans SC", "PingFang SC", "Microsoft YaHei", sans-serif;
}

/* 步骤卡片 */
.step-card {
    background: #f8f9fc;
    border-radius: 14px;
    padding: 1.5rem 2rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
}

/* 知识标题 */
.know-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #2c5f8a;
    margin-bottom: 0.4rem;
}

/* 知识节标题 */
.section-badge {
    display: inline-block;
    background: #2c5f8a;
    color: white;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 0.8rem;
    font-weight: 600;
    margin-bottom: 0.5rem;
}

/* 知识正文 */
.know-body {
    font-size: 1.05rem;
    line-height: 1.9;
    color: #2d2d2d;
    white-space: pre-wrap;
}

/* 进度条文字 */
.progress-label {
    font-size: 0.85rem;
    color: #888;
    margin-bottom: 4px;
}

/* 答对/答错提示 */
.feedback-correct {
    background: #d4edda;
    border-left: 4px solid #28a745;
    padding: 0.6rem 1rem;
    border-radius: 6px;
    color: #155724;
    font-weight: 600;
}
.feedback-wrong {
    background: #f8d7da;
    border-left: 4px solid #dc3545;
    padding: 0.6rem 1rem;
    border-radius: 6px;
    color: #721c24;
    font-weight: 600;
}

/* 单选题选项按钮 */
div[data-testid="stRadio"] label {
    font-size: 1.0rem;
    padding: 4px 0;
}

/* 侧边栏节选择器 */
.section-select label {
    font-size: 0.85rem !important;
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# Session State 初始化
# ──────────────────────────────────────────────
def init_state():
    defaults = {
        "items": [],          # 解析后的全部条目
        "knowledge": [],      # 仅知识条目
        "questions": [],      # 仅题目条目
        "sections": [],       # 所有节名
        "current_section": "全部",  # 当前选中节
        "filtered_knowledge": [],  # 当前节的知识
        "filtered_questions": [],  # 当前节的题目
        "step": 1,            # 当前步骤 1/2/3
        "know_idx": 0,        # 知识展示索引
        "fill_idx": 0,        # 填空索引
        "quiz_idx": 0,        # 练习题索引
        "quiz_score": 0,
        "quiz_total": 0,
        "quiz_answered": False,
        "quiz_selected": None,
        "quiz_order": [],     # 随机打乱的题目顺序
        "fill_answers": {},   # {blank_key: user_answer}
        "fill_submitted": False,
        "file_loaded": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ──────────────────────────────────────────────
# 工具函数
# ──────────────────────────────────────────────

def load_items(items: list):
    st.session_state.items = items
    st.session_state.knowledge = [x for x in items if x['type'] == 'knowledge']
    st.session_state.questions = [x for x in items if x['type'] == 'question']

    # 提取所有节
    sections = []
    for item in items:
        s = item.get('section', '')
        if s and s not in sections:
            sections.append(s)
    st.session_state.sections = sections

    st.session_state.current_section = "全部"
    _apply_section_filter()

    st.session_state.know_idx = 0
    st.session_state.fill_idx = 0
    st.session_state.quiz_idx = 0
    st.session_state.quiz_score = 0
    st.session_state.quiz_total = 0
    st.session_state.quiz_answered = False
    st.session_state.quiz_selected = None
    st.session_state.quiz_order = list(range(len(st.session_state.filtered_questions)))
    random.shuffle(st.session_state.quiz_order)
    st.session_state.fill_answers = {}
    st.session_state.fill_submitted = False
    st.session_state.file_loaded = True


def _apply_section_filter():
    sec = st.session_state.current_section
    if sec == "全部":
        st.session_state.filtered_knowledge = st.session_state.knowledge
        st.session_state.filtered_questions = st.session_state.questions
    else:
        st.session_state.filtered_knowledge = [x for x in st.session_state.knowledge if x.get('section') == sec]
        st.session_state.filtered_questions = [x for x in st.session_state.questions if x.get('section') == sec]


def _get_knowledge_with_blanks(knowledge_list):
    """只返回有填空词的知识条目"""
    return [x for x in knowledge_list if x.get('blanks')]


def render_blank_text(template: str, answers: dict, submitted: bool, correct_blanks: list) -> str:
    """把 [BLANK_N] 替换成用户填写的内容（已提交则标色）"""
    def replacer(m):
        n = int(m.group(1))
        key = f"blank_{n}"
        user_val = answers.get(key, "")
        if not submitted:
            underline = "─" * max(4, len(correct_blanks[n-1]) * 2 + 2) if n-1 < len(correct_blanks) else "────────"
            filled = f'<u style="color:#2c5f8a;font-weight:600">{user_val}</u>' if user_val else f'<span style="color:#aaa">{underline}</span>'
        else:
            correct = correct_blanks[n-1] if n-1 < len(correct_blanks) else ""
            if user_val == correct:
                filled = f'<span style="background:#d4edda;color:#155724;padding:1px 4px;border-radius:4px;font-weight:700">{user_val}</span>'
            else:
                filled = (f'<span style="background:#f8d7da;color:#721c24;padding:1px 4px;border-radius:4px;text-decoration:line-through">{user_val}</span>'
                          f' <span style="background:#d4edda;color:#155724;padding:1px 4px;border-radius:4px;font-weight:700">{correct}</span>')
        return filled
    return re.sub(r'\[BLANK_(\d+)\]', replacer, template)


# ──────────────────────────────────────────────
# 侧边栏
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📚 背记系统")
    st.markdown("---")

    # 文件上传
    uploaded = st.file_uploader(
        "上传 Word 文档（.docx）",
        type=["docx"],
        help="支持含知识点和选择题的 Word 文档",
    )
    if uploaded:
        try:
            with st.spinner("解析文档中…"):
                items = parse_docx(uploaded)
            if items:
                load_items(items)
                st.success(f"✅ 解析完成：{len(st.session_state.knowledge)} 条知识，{len(st.session_state.questions)} 道题")
            else:
                st.warning("未识别到有效内容，已加载演示数据")
                load_items(DEMO_ITEMS)
        except Exception as e:
            st.error(f"解析失败：{e}")
            load_items(DEMO_ITEMS)
    elif not st.session_state.file_loaded:
        if st.button("📋 加载演示数据"):
            load_items(DEMO_ITEMS)

    st.markdown("---")

    # 节选择器 + 步骤导航
    if st.session_state.file_loaded:
        # 节选择
        sections = st.session_state.sections
        section_options = ["全部"] + sections
        selected_section = st.selectbox(
            "📂 选择章节",
            options=section_options,
            index=0,
            key="section_selector",
        )
        if selected_section != st.session_state.current_section:
            st.session_state.current_section = selected_section
            _apply_section_filter()
            st.session_state.know_idx = 0
            st.session_state.fill_idx = 0
            st.session_state.quiz_idx = 0
            st.session_state.quiz_score = 0
            st.session_state.quiz_total = 0
            st.session_state.quiz_answered = False
            st.session_state.quiz_selected = None
            st.session_state.quiz_order = list(range(len(st.session_state.filtered_questions)))
            random.shuffle(st.session_state.quiz_order)
            st.session_state.fill_answers = {}
            st.session_state.fill_submitted = False

        # 统计
        fk = st.session_state.filtered_knowledge
        fq = st.session_state.filtered_questions
        fkb = _get_knowledge_with_blanks(fk)
        st.caption(f"当前节：{len(fk)} 条知识 · {len(fkb)} 条有填空 · {len(fq)} 道题")

        st.markdown("---")
        st.markdown("### 学习步骤")
        for step_no, (icon, label) in enumerate([
            ("👁️", "知识展示"),
            ("🫧", "气泡池填空"),
            ("⚡", "快速练习"),
        ], start=1):
            active = st.session_state.step == step_no
            if st.button(
                f"{icon} {step_no}. {label}",
                key=f"nav_step_{step_no}",
                use_container_width=True,
                type="primary" if active else "secondary",
            ):
                st.session_state.step = step_no
                st.rerun()

    st.markdown("---")

    # DeepSeek API 配置
    import parser as _parser
    st.markdown("#### 🤖 AI 干扰项生成")
    current_key = _parser.DEEPSEEK_API_KEY
    api_key_input = st.text_input(
        "DeepSeek API Key",
        value=current_key,
        type="password",
        help="填入 DeepSeek API Key 后，填空练习的干扰词将由 AI 智能生成，\n否则仅从文档内已有填空词中抽取。",
        key="deepseek_api_key_input",
    )
    if api_key_input != current_key:
        _parser.DEEPSEEK_API_KEY = api_key_input
        st.caption("✅ API Key 已更新" if api_key_input else "⚠️ 未配置 API Key，使用文档内抽取模式")

    st.markdown("---")
    st.markdown(
        "<small>📌 格式说明：Word文档中<br>"
        "• <b>知识段落</b>：普通正文<br>"
        "• <b>填空词</b>：下划线格式的词<br>"
        "• <b>选择题</b>：【题目内容】+【正确选项】+【错误选项】<br>"
        "• 兼容 A./B./C./D. 传统格式</small>",
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────
# 主内容区
# ──────────────────────────────────────────────
if not st.session_state.file_loaded:
    # 欢迎页
    st.markdown("""
    <div style="text-align:center; padding: 4rem 2rem;">
        <div style="font-size:4rem">📚</div>
        <h1 style="color:#2c5f8a; margin-top:0.5rem">背记系统</h1>
        <p style="font-size:1.1rem; color:#555; max-width:500px; margin:auto">
            上传包含知识点和选择题的 Word 文档，<br>
            系统自动解析并生成三步学习界面。
        </p>
        <br>
        <div style="display:flex; justify-content:center; gap:2rem; flex-wrap:wrap">
            <div style="background:#e8eef5;padding:1.2rem 1.8rem;border-radius:12px;text-align:center;width:160px">
                <div style="font-size:2.2rem">👁️</div>
                <b>知识展示</b><br><small>记忆原文</small>
            </div>
            <div style="background:#e8eef5;padding:1.2rem 1.8rem;border-radius:12px;text-align:center;width:160px">
                <div style="font-size:2.2rem">🫧</div>
                <b>气泡池填空</b><br><small>拖拽填词</small>
            </div>
            <div style="background:#e8eef5;padding:1.2rem 1.8rem;border-radius:12px;text-align:center;width:160px">
                <div style="font-size:2.2rem">⚡</div>
                <b>快速练习</b><br><small>单选刷题</small>
            </div>
        </div>
        <br>
        <p style="color:#aaa; font-size:0.9rem">← 在左侧上传文档，或点击"加载演示数据"开始体验</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════
#  STEP 1  知识展示
# ══════════════════════════════════════════════
if st.session_state.step == 1:
    knows = st.session_state.filtered_knowledge
    if not knows:
        st.info("当前章节无知识段落。")
        st.stop()

    total = len(knows)
    idx = min(st.session_state.know_idx, total - 1)
    st.session_state.know_idx = idx
    item = knows[idx]

    # 标题栏
    col_title, col_prog = st.columns([3, 1])
    with col_title:
        st.markdown("## 👁️ 第一步：知识展示")
    with col_prog:
        st.markdown(f'<p class="progress-label">知识点 {idx+1} / {total}</p>', unsafe_allow_html=True)
        st.progress((idx + 1) / total)

    # 知识卡片
    section_html = f'<div class="section-badge">{item.get("section", "")}</div>' if item.get("section") else ""
    title_html = f'<div class="know-title">{item["title"]}</div>' if item.get("title") else ""
    body_text = item.get("content", "")
    # 把 [BLANK_N] 替换回答案词，让展示时看到完整原文
    blanks = item.get("blanks", [])
    display_text = body_text
    for bi, b in enumerate(blanks, 1):
        display_text = display_text.replace(f"[BLANK_{bi}]", f'<u style="color:#e07b00;font-weight:700">{b}</u>')

    st.markdown(f"""
    <div class="step-card">
        {section_html}
        {title_html}
        <div class="know-body">{display_text}</div>
    </div>
    """, unsafe_allow_html=True)

    # 导航按钮
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        if idx > 0:
            if st.button("◀ 上一条", use_container_width=True):
                st.session_state.know_idx -= 1
                st.rerun()
    with c2:
        if st.button("✅ 完成，进入填空练习 →", type="primary", use_container_width=True):
            st.session_state.step = 2
            st.session_state.fill_idx = 0
            st.session_state.fill_answers = {}
            st.session_state.fill_submitted = False
            st.rerun()
    with c3:
        if idx < total - 1:
            if st.button("下一条 ▶", use_container_width=True):
                st.session_state.know_idx += 1
                st.rerun()


# ══════════════════════════════════════════════
#  STEP 2  气泡池填空
# ══════════════════════════════════════════════
elif st.session_state.step == 2:
    all_knows = st.session_state.filtered_knowledge
    knows_with_blanks = _get_knowledge_with_blanks(all_knows)

    if not knows_with_blanks:
        st.info("当前章节无带填空的知识段落。")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("↩️ 回到知识展示"):
                st.session_state.step = 1
                st.rerun()
        with c2:
            if st.button("⚡ 进入快速练习"):
                st.session_state.step = 3
                st.rerun()
        st.stop()

    total = len(knows_with_blanks)
    idx = min(st.session_state.fill_idx, total - 1)
    st.session_state.fill_idx = idx
    item = knows_with_blanks[idx]
    blanks = item.get("blanks", [])
    template = item.get("template", item.get("content", ""))

    col_title, col_prog = st.columns([3, 1])
    with col_title:
        st.markdown("## 🫧 第二步：气泡池填空")
    with col_prog:
        st.markdown(f'<p class="progress-label">填空 {idx+1} / {total}</p>', unsafe_allow_html=True)
        st.progress((idx + 1) / total)

    submitted = st.session_state.fill_submitted
    answers = st.session_state.fill_answers

    # 生成选项池（正确词 + 干扰词，打乱）
    items_safe = st.session_state.items if st.session_state.items else []
    distractors = generate_distractors(items_safe, blanks, n=max(3, len(blanks)))
    option_pool = blanks + distractors
    random.seed(f"fill_{idx}_{len(blanks)}")   # 固定打乱顺序（同一题不变）
    random.shuffle(option_pool)

    # ── 渲染填空文本 ──
    section_html = f'<div class="section-badge">{item.get("section", "")}</div>' if item.get("section") else ""
    title_html = f'<div class="know-title">{item["title"]}</div>' if item.get("title") else ""
    st.markdown(f'<div class="step-card">{section_html}{title_html}', unsafe_allow_html=True)

    if submitted:
        result_html = render_blank_text(template, answers, submitted=True, correct_blanks=blanks)
        st.markdown(f'<div class="know-body">{result_html}</div>', unsafe_allow_html=True)
    else:
        st.markdown(
            '<div class="know-body" style="margin-bottom:0.5rem">'
            + render_blank_text(template, answers, submitted=False, correct_blanks=blanks)
            + '</div>',
            unsafe_allow_html=True
        )

    st.markdown('</div>', unsafe_allow_html=True)

    if not submitted:
        # 气泡池
        st.markdown("#### 🫧 词语气泡池（点击选词填入各空格）")

        blank_count = len(blanks)
        cols_per_row = min(4, blank_count)
        cols = st.columns(cols_per_row)
        for bi in range(blank_count):
            n = bi + 1
            key = f"blank_{n}"
            with cols[bi % cols_per_row]:
                current_val = answers.get(key, "")
                opts = ["（请选择）"] + option_pool
                default_idx = 0
                if current_val and current_val in option_pool:
                    default_idx = opts.index(current_val)
                chosen = st.selectbox(
                    f"第 {n} 空",
                    options=opts,
                    index=default_idx,
                    key=f"sel_blank_{idx}_{n}",
                )
                if chosen != "（请选择）":
                    st.session_state.fill_answers[key] = chosen

        st.markdown("---")
        c1, c2, c3 = st.columns([1, 2, 1])
        with c1:
            if st.button("◀ 上一条", disabled=(idx == 0)):
                st.session_state.fill_idx -= 1
                st.session_state.fill_answers = {}
                st.session_state.fill_submitted = False
                st.rerun()
        with c2:
            all_filled = all(f"blank_{n+1}" in st.session_state.fill_answers for n in range(blank_count))
            if st.button("✅ 提交答案", type="primary", use_container_width=True, disabled=not all_filled):
                st.session_state.fill_submitted = True
                st.rerun()
        with c3:
            if st.button("跳过 →"):
                if idx < total - 1:
                    st.session_state.fill_idx += 1
                    st.session_state.fill_answers = {}
                    st.session_state.fill_submitted = False
                else:
                    st.session_state.step = 3
                st.rerun()
    else:
        # 提交后统计
        correct_count = sum(
            1 for n in range(1, len(blanks)+1)
            if answers.get(f"blank_{n}", "") == blanks[n-1]
        )
        total_blanks = len(blanks)

        if correct_count == total_blanks:
            st.markdown('<div class="feedback-correct">🎉 全部正确！</div>', unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div class="feedback-wrong">⚠️ 答对 {correct_count} / {total_blanks} 个</div>',
                unsafe_allow_html=True
            )

        st.markdown("")
        c1, c2, c3 = st.columns([1, 2, 1])
        with c1:
            if st.button("◀ 上一条", disabled=(idx == 0)):
                st.session_state.fill_idx -= 1
                st.session_state.fill_answers = {}
                st.session_state.fill_submitted = False
                st.rerun()
        with c2:
            if st.button("🔄 重新练习", use_container_width=True):
                st.session_state.fill_answers = {}
                st.session_state.fill_submitted = False
                st.rerun()
        with c3:
            label = "下一条 ▶" if idx < total - 1 else "进入快速练习 ⚡"
            if st.button(label, type="primary", use_container_width=True):
                if idx < total - 1:
                    st.session_state.fill_idx += 1
                    st.session_state.fill_answers = {}
                    st.session_state.fill_submitted = False
                else:
                    st.session_state.step = 3
                    st.session_state.quiz_idx = 0
                    st.session_state.quiz_score = 0
                    st.session_state.quiz_total = 0
                    st.session_state.quiz_answered = False
                    st.session_state.quiz_selected = None
                st.rerun()


# ══════════════════════════════════════════════
#  STEP 3  快速练习
# ══════════════════════════════════════════════
elif st.session_state.step == 3:
    questions = st.session_state.filtered_questions
    if not questions:
        st.info("当前章节无选择题。")
        if st.button("↩️ 回到知识展示"):
            st.session_state.step = 1
            st.rerun()
        st.stop()

    order = st.session_state.quiz_order
    total_q = len(order)
    qi = st.session_state.quiz_idx

    # 标题 + 分数
    col_t, col_s = st.columns([3, 1])
    with col_t:
        st.markdown("## ⚡ 第三步：快速练习")
    with col_s:
        score = st.session_state.quiz_score
        done = st.session_state.quiz_total
        st.metric("得分", f"{score} / {done}", delta=None)

    if qi >= total_q:
        # 全部完成
        pct = int(score / total_q * 100) if total_q else 0
        st.markdown(f"""
        <div class="step-card" style="text-align:center">
            <div style="font-size:3rem">{'🏆' if pct >= 80 else '📝'}</div>
            <h2 style="color:#2c5f8a">练习完成！</h2>
            <p style="font-size:1.3rem">得分：<b>{score} / {total_q}</b>（{pct}%）</p>
            <p style="color:#555">{'太棒了，全部掌握！' if pct == 100 else '再接再厉，继续加油！' if pct >= 60 else '建议返回知识展示再复习一遍～'}</p>
        </div>
        """, unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 重新练习（随机顺序）", use_container_width=True, type="primary"):
                new_order = list(range(len(questions)))
                random.shuffle(new_order)
                st.session_state.quiz_order = new_order
                st.session_state.quiz_idx = 0
                st.session_state.quiz_score = 0
                st.session_state.quiz_total = 0
                st.session_state.quiz_answered = False
                st.session_state.quiz_selected = None
                st.rerun()
        with c2:
            if st.button("↩️ 回到知识展示", use_container_width=True):
                st.session_state.step = 1
                st.session_state.know_idx = 0
                st.rerun()
        st.stop()

    # 当前题目
    q_idx = order[qi]
    q = questions[q_idx]
    answered = st.session_state.quiz_answered
    selected = st.session_state.quiz_selected
    correct_answer = q.get("answer", "")

    # 题目卡片
    section_badge = f'<div class="section-badge">{q.get("section", "")}</div>' if q.get("section") else ""
    st.markdown(f"""
    <div class="step-card">
        {section_badge}
        <div style="color:#888;font-size:0.85rem;margin-bottom:0.4rem">题目 {qi+1} / {total_q}</div>
        <div style="font-size:1.1rem;font-weight:600;color:#1a1a1a;line-height:1.7">{q['stem']}</div>
    </div>
    """, unsafe_allow_html=True)

    st.progress((qi + 1) / total_q)

    # 选项
    opt_labels = sorted(q['options'].keys())
    opt_texts = [f"{k}. {q['options'][k]}" for k in opt_labels]

    if not answered:
        chosen = st.radio(
            "请选择答案：",
            options=opt_texts,
            index=None,
            key=f"quiz_radio_{qi}",
            label_visibility="collapsed",
        )
        st.markdown("")
        if st.button("✅ 确认答案", type="primary", disabled=(chosen is None)):
            chosen_letter = chosen[0] if chosen else ""
            st.session_state.quiz_selected = chosen_letter
            st.session_state.quiz_answered = True
            st.session_state.quiz_total += 1
            if chosen_letter == correct_answer:
                st.session_state.quiz_score += 1
            st.rerun()
    else:
        # 显示所有选项，标色
        for k in opt_labels:
            text = q['options'][k]
            if k == correct_answer and k == selected:
                color = "#d4edda"; icon = "✅"; weight = "700"
            elif k == correct_answer:
                color = "#d4edda"; icon = "✅"; weight = "700"
            elif k == selected:
                color = "#f8d7da"; icon = "❌"; weight = "600"
            else:
                color = "#f8f9fc"; icon = ""; weight = "400"
            st.markdown(
                f'<div style="background:{color};padding:0.55rem 1rem;border-radius:8px;'
                f'margin:0.25rem 0;font-weight:{weight}">{icon} {k}. {text}</div>',
                unsafe_allow_html=True
            )

        st.markdown("")
        if selected == correct_answer:
            st.markdown('<div class="feedback-correct">🎉 回答正确！</div>', unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div class="feedback-wrong">❌ 回答错误，正确答案是 <b>{correct_answer}</b></div>',
                unsafe_allow_html=True
            )

        st.markdown("")
        c1, c2 = st.columns([1, 3])
        with c2:
            label = "下一题 ▶" if qi < total_q - 1 else "查看成绩 🏆"
            if st.button(label, type="primary", use_container_width=True):
                st.session_state.quiz_idx += 1
                st.session_state.quiz_answered = False
                st.session_state.quiz_selected = None
                st.rerun()
