import streamlit as st
import anthropic
from docx import Document
from docx.shared import RGBColor, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import json
from datetime import datetime

# ─────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="GEO-Scan | 크림웍스",
    layout="wide",
    initial_sidebar_state="collapsed"
)

CW_PURPLE = "#6B4EFF"
CW_PURPLE_LIGHT = "#EDE9FF"
CW_PURPLE_RGB = (107, 78, 255)

st.markdown(f"""
<style>
  .main .block-container {{ padding-top: 1.5rem; max-width: 860px; margin: 0 auto; }}
  .brand-header {{
      padding: 24px 28px;
      border-radius: 14px;
      margin-bottom: 1.5rem;
      color: white;
  }}
  .step-label {{
      font-size: 0.72rem;
      font-weight: 600;
      color: #999;
      letter-spacing: 1px;
      text-transform: uppercase;
      margin-bottom: 2px;
  }}
  .step-title {{
      font-size: 1.1rem;
      font-weight: 700;
      color: #1a1a1a;
      margin-bottom: 0.8rem;
  }}
  .field-label {{
      font-size: 0.82rem;
      font-weight: 600;
      color: #444;
      margin-bottom: 4px;
      margin-top: 10px;
  }}
  .required {{ color: #e74c3c; }}
  .cw-box {{
      background: {CW_PURPLE_LIGHT};
      border-left: 4px solid {CW_PURPLE};
      padding: 14px 18px;
      border-radius: 0 10px 10px 0;
      margin: 8px 0 16px 0;
      font-size: 0.88rem;
      color: #3a2d7a;
  }}
  .info-box {{
      background: #f0f4ff;
      border: 1px solid #c7d3ff;
      border-radius: 10px;
      padding: 14px 18px;
      font-size: 0.88rem;
      color: #2c3e80;
      margin-bottom: 1rem;
  }}
  .mention-yes {{
      background: #d4edda; color: #155724;
      padding: 3px 10px; border-radius: 10px;
      font-size: 0.78rem; font-weight: 700;
  }}
  .mention-no {{
      background: #f8d7da; color: #721c24;
      padding: 3px 10px; border-radius: 10px;
      font-size: 0.78rem; font-weight: 700;
  }}
  .q-card {{
      border: 1px solid #e8e8e8;
      border-radius: 10px;
      padding: 16px 18px;
      margin-bottom: 12px;
      background: white;
  }}
  .q-num {{
      display: inline-block;
      background: #111;
      color: white;
      width: 22px; height: 22px;
      border-radius: 50%;
      text-align: center;
      line-height: 22px;
      font-size: 0.75rem;
      font-weight: 700;
      margin-right: 8px;
  }}
  .divider {{ border: none; border-top: 1px solid #eee; margin: 1.5rem 0; }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 세션 초기화
# ─────────────────────────────────────────────
def init():
    defaults = {
        'step': 1,
        'api_key': '',
        'brand_name': '',
        'brand_color': '#000000',
        'brand_category': '',
        'brand_usp': '',
        'brand_target': '',
        'brand_competitors': '',
        'brand_negative': '',
        'brand_focus': '',
        'questions': [],
        'questions_confirmed': False,
        'answers': {
            'off': {i: '' for i in range(1, 8)},
            'on':  {i: '' for i in range(1, 8)},
            'gem': {i: '' for i in range(1, 8)},
        },
        'analysis_result': '',
        'cw_insights': [''] * 7,
        'overall_diagnosis': '',
        'priority_actions': '🚨 즉시 실행:\n\n⭐ 1개월 내:\n\n3개월 내:',
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init()

# ─────────────────────────────────────────────
# 헬퍼
# ─────────────────────────────────────────────
def hex_to_rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.replace('#', ''))
    tcPr.append(shd)

def check_mention(text, brand_name):
    if not text.strip():
        return None
    keywords = [brand_name, brand_name[:2], brand_name.replace(' ', '')]
    return any(kw in text for kw in keywords)

def get_client():
    return anthropic.Anthropic(api_key=st.session_state.api_key)

# ─────────────────────────────────────────────
# 헤더
# ─────────────────────────────────────────────
bc = st.session_state.brand_color
bn = st.session_state.brand_name or 'GEO-Scan'
st.markdown(f"""
<div class="brand-header" style="background: linear-gradient(135deg, #0f0f0f 60%, {bc}99);">
  <div style="font-size:0.78rem;color:#888;letter-spacing:2px;font-weight:600;">CREAMWORKS  ·  GEO-Scan v2.0</div>
  <div style="font-size:1.7rem;font-weight:800;margin:6px 0 4px;">{bn} AI 진단 시스템</div>
  <div style="font-size:0.85rem;color:#bbb;">브랜드 AI 검색 노출 현황 진단 · GEO 전략 제안서 자동 생성</div>
</div>
""", unsafe_allow_html=True)

# 스텝 인디케이터
step_names = ["브랜드 입력", "질문 확인", "답변 수집", "분석·인사이트", "보고서"]
cols = st.columns(5)
for i, (col, name) in enumerate(zip(cols, step_names)):
    n = i + 1
    active = st.session_state.step == n
    done = st.session_state.step > n
    bg = CW_PURPLE if active else ('#27ae60' if done else '#ddd')
    tc = 'white'
    label_c = '#1a1a1a' if active else ('#27ae60' if done else '#aaa')
    col.markdown(f"""
    <div style="text-align:center;padding:4px 0">
      <div style="width:26px;height:26px;border-radius:50%;background:{bg};
                  color:{tc};display:inline-flex;align-items:center;justify-content:center;
                  font-size:0.72rem;font-weight:700;margin-bottom:3px">
        {'✓' if done else n}
      </div>
      <div style="font-size:0.7rem;color:{label_c};font-weight:{'700' if active else '400'}">{name}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<hr class='divider'>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# STEP 1: 브랜드 정보 입력
# ─────────────────────────────────────────────
if st.session_state.step == 1:
    st.markdown('<div class="step-label">STEP 1</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-title">브랜드 정보 입력</div>', unsafe_allow_html=True)

    st.markdown('<div class="cw-box">💜 많이 입력할수록 질문 퀄리티가 좋아집니다. 핵심 USP와 강조 포인트는 꼭 입력해주세요.</div>', unsafe_allow_html=True)

    # API 키
    st.markdown('<div class="field-label">🔑 Anthropic API Key <span class="required">*</span></div>', unsafe_allow_html=True)
    api_key = st.text_input("api", value=st.session_state.api_key,
                             type="password", placeholder="sk-ant-api03-...",
                             label_visibility="collapsed")
    st.session_state.api_key = api_key
    st.caption("API 키는 브라우저 세션에만 유지되며 저장되지 않습니다.")

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="field-label">브랜드명 <span class="required">*</span></div>', unsafe_allow_html=True)
        st.session_state.brand_name = st.text_input(
            "브랜드명", value=st.session_state.brand_name,
            placeholder="예: 교촌치킨", label_visibility="collapsed")

        st.markdown('<div class="field-label">카테고리 <span class="required">*</span></div>', unsafe_allow_html=True)
        st.session_state.brand_category = st.text_input(
            "카테고리", value=st.session_state.brand_category,
            placeholder="예: 치킨 프랜차이즈", label_visibility="collapsed")

        st.markdown('<div class="field-label">브랜드 공식 컬러</div>', unsafe_allow_html=True)
        st.session_state.brand_color = st.color_picker(
            "컬러", value=st.session_state.brand_color, label_visibility="collapsed")
        bc_preview = st.session_state.brand_color
        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:10px;margin-top:6px;margin-bottom:2px">
          <div style="width:32px;height:32px;background:{bc_preview};
                      border-radius:6px;border:1px solid #ddd;"></div>
          <span style="font-size:0.82rem;color:#666">{bc_preview} — 보고서·질문지에 자동 적용됩니다</span>
        </div>
        <div style="font-size:0.78rem;color:#aaa;margin-bottom:8px">
          모르시면 구글에 "{st.session_state.brand_name or '브랜드명'} 브랜드 컬러" 검색
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="field-label">경쟁 브랜드</div>', unsafe_allow_html=True)
        st.session_state.brand_competitors = st.text_input(
            "경쟁사", value=st.session_state.brand_competitors,
            placeholder="예: BBQ, BHC", label_visibility="collapsed")

        st.markdown('<div class="field-label">주요 타겟 소비자</div>', unsafe_allow_html=True)
        st.session_state.brand_target = st.text_input(
            "타겟", value=st.session_state.brand_target,
            placeholder="예: 남녀노소, 배달앱 구매 중심", label_visibility="collapsed")

    with col2:
        st.markdown('<div class="field-label">핵심 USP (차별점) <span class="required">*</span></div>', unsafe_allow_html=True)
        st.session_state.brand_usp = st.text_area(
            "USP", value=st.session_state.brand_usp, height=110,
            placeholder="예: 간장치킨 원조, 35년 업력, 비밀 마늘간장 소스, 붓질 공정, 국내산 재료",
            label_visibility="collapsed")

        st.markdown('<div class="field-label">부정 이미지 / 약점</div>', unsafe_allow_html=True)
        st.session_state.brand_negative = st.text_area(
            "약점", value=st.session_state.brand_negative, height=80,
            placeholder="예: 가격 인상 선도, 배달 이중가격 논란",
            label_visibility="collapsed")

        st.markdown('<div class="field-label">지금 강조하고 싶은 포인트 <span class="required">*</span></div>', unsafe_allow_html=True)
        st.session_state.brand_focus = st.text_area(
            "포인트", value=st.session_state.brand_focus, height=80,
            placeholder="예: 소비자가 교촌을 선택하는 이유 부각, 앱 구매 중심 GEO 세팅 방향",
            label_visibility="collapsed")

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    if st.button("🔍  질문 7개 생성하기", type="primary", use_container_width=True):
        # 유효성 검사
        if not st.session_state.api_key:
            st.error("Anthropic API Key를 입력해주세요.")
        elif not st.session_state.brand_name:
            st.error("브랜드명을 입력해주세요.")
        elif not st.session_state.brand_category:
            st.error("카테고리를 입력해주세요.")
        elif not st.session_state.brand_usp:
            st.error("핵심 USP를 입력해주세요.")
        else:
            with st.spinner("Claude가 브랜드를 분석하고 질문을 설계하는 중입니다..."):
                try:
                    client = get_client()

                    prompt = f"""당신은 GEO(Generative Engine Optimization) 전문가입니다.
아래 브랜드 정보를 분석해서, 소비자가 ChatGPT·Gemini에 실제로 물어볼 법한 GEO 진단 질문 7개를 설계해주세요.

브랜드 정보:
- 브랜드명: {st.session_state.brand_name}
- 카테고리: {st.session_state.brand_category}
- 핵심 USP: {st.session_state.brand_usp}
- 주요 타겟: {st.session_state.brand_target}
- 경쟁 브랜드: {st.session_state.brand_competitors}
- 부정 이미지/약점: {st.session_state.brand_negative}
- 강조 포인트: {st.session_state.brand_focus}

질문 설계 원칙:
1. 브랜드명이 절대 들어가면 안 됨 (브랜드를 모르는 소비자가 AI에게 묻는 질문)
2. 실제 소비자가 쓰는 구어체
3. AIJ 5단계 커버: DISCOVER(2개) / CONSIDER(3개) / DECIDE(2개)
4. 브랜드 USP와 직접 연결되는 질문 우선
5. 부정 이미지 방어 질문 1개 이상 포함

반드시 아래 JSON 형식으로만 응답 (다른 텍스트 없이):
{{
  "questions": [
    {{
      "question": "질문 내용",
      "stage": "DISCOVER 또는 CONSIDER 또는 DECIDE",
      "type": "유형명 (예: 카테고리 진입 — 브랜드 선택 첫 질문)",
      "check_point": "확인 포인트: {st.session_state.brand_name}이 어떤 맥락에서 등장하는지 + 경쟁사 대비 포지션",
      "data": [
        {{"source": "출처기관", "content": "구체적 데이터 내용", "year": "2024"}},
        {{"source": "출처기관", "content": "구체적 데이터 내용", "year": "2025"}},
        {{"source": "출처기관", "content": "구체적 데이터 내용", "year": "2025"}}
      ]
    }}
  ]
}}"""

                    message = client.messages.create(
                        model="claude-sonnet-4-20250514",
                        max_tokens=4000,
                        messages=[{"role": "user", "content": prompt}]
                    )

                    raw = message.content[0].text.strip()
                    if "```" in raw:
                        raw = raw.split("```")[1]
                        if raw.startswith("json"):
                            raw = raw[4:]
                    raw = raw.strip()

                    data = json.loads(raw)
                    st.session_state.questions = data['questions']
                    st.session_state.questions_confirmed = False
                    st.session_state.step = 2
                    st.rerun()

                except json.JSONDecodeError:
                    st.error("질문 생성 중 파싱 오류가 발생했습니다. 다시 시도해주세요.")
                except Exception as e:
                    st.error(f"오류 발생: {str(e)}")

# ─────────────────────────────────────────────
# STEP 2: 질문 확인 및 수정
# ─────────────────────────────────────────────
elif st.session_state.step == 2:
    st.markdown('<div class="step-label">STEP 2</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-title">진단 질문 7개 확인 및 수정</div>', unsafe_allow_html=True)
    st.markdown('<div class="cw-box">💜 Claude가 설계한 질문입니다. 수정이 필요하면 직접 편집하세요. 브랜드명이 포함된 질문은 사용하지 않습니다.</div>', unsafe_allow_html=True)

    updated_qs = []
    for i, q_data in enumerate(st.session_state.questions):
        n = i + 1
        stage = q_data.get('stage', '')
        qtype = q_data.get('type', '')
        check = q_data.get('check_point', '')
        data_list = q_data.get('data', [])

        with st.expander(f"Q{n} · [{stage}] {qtype}", expanded=(i < 2)):
            new_q = st.text_input(
                f"질문 내용",
                value=q_data.get('question', ''),
                key=f"q_edit_{i}",
                label_visibility="visible"
            )
            st.caption(f"**확인 포인트:** {check}")

            if data_list:
                st.markdown("**📊 선정 근거 데이터**")
                cols = st.columns(3)
                for j, d in enumerate(data_list[:3]):
                    with cols[j]:
                        st.markdown(f"""
                        <div style="background:#f8f9fa;border-radius:8px;padding:10px;font-size:0.78rem;">
                          <div style="font-weight:700;color:#555;margin-bottom:4px">{d.get('source','')} ({d.get('year','')})</div>
                          <div style="color:#333">{d.get('content','')}</div>
                        </div>""", unsafe_allow_html=True)

            q_updated = dict(q_data)
            q_updated['question'] = new_q
            updated_qs.append(q_updated)

    col_back, col_confirm = st.columns([1, 3])
    with col_back:
        if st.button("← 다시 입력", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
    with col_confirm:
        if st.button("✅ 질문 확정 — 답변 수집으로 이동", type="primary", use_container_width=True):
            st.session_state.questions = updated_qs
            st.session_state.questions_confirmed = True
            st.session_state.answers = {
                'off': {i: '' for i in range(1, 8)},
                'on':  {i: '' for i in range(1, 8)},
                'gem': {i: '' for i in range(1, 8)},
            }
            st.session_state.step = 3
            st.rerun()

# ─────────────────────────────────────────────
# STEP 3: 답변 수집
# ─────────────────────────────────────────────
elif st.session_state.step == 3:
    st.markdown('<div class="step-label">STEP 3</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-title">AI 답변 수집</div>', unsafe_allow_html=True)

    # 진행 현황
    off_cnt = sum(1 for v in st.session_state.answers['off'].values() if v.strip())
    on_cnt  = sum(1 for v in st.session_state.answers['on'].values() if v.strip())
    gem_cnt = sum(1 for v in st.session_state.answers['gem'].values() if v.strip())
    total   = off_cnt + on_cnt + gem_cnt

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("GPT 검색OFF", f"{off_cnt}/7")
    c2.metric("GPT 검색ON", f"{on_cnt}/7")
    c3.metric("Gemini", f"{gem_cnt}/7")
    c4.metric("전체", f"{total}/21")
    st.progress(total / 21)

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    ai_tabs = st.tabs(["🔵 ChatGPT 검색OFF", "🟢 ChatGPT 검색ON", "🟠 Gemini"])
    ai_keys   = ['off', 'on', 'gem']
    ai_labels = ['ChatGPT 검색OFF', 'ChatGPT 검색ON', 'Gemini']
    ai_instrs = [
        "ChatGPT → 새 채팅 → 메모리 OFF + 검색 OFF → Q1~Q7 순서대로 동일 채팅 입력 → 각 답변 복사 후 붙여넣기",
        "ChatGPT → 새 채팅 → 메모리 OFF + 검색 ON → Q1~Q7 순서대로 동일 채팅 입력 → 각 답변 복사 후 붙여넣기",
        "시크릿 모드 → gemini.google.com (로그아웃) → 새 채팅 → Q1~Q7 순서대로 입력 → 각 답변 복사 후 붙여넣기"
    ]

    for ai_tab, key, label, instr in zip(ai_tabs, ai_keys, ai_labels, ai_instrs):
        with ai_tab:
            st.markdown(f'<div class="info-box">📌 {instr}</div>', unsafe_allow_html=True)

            for i, q_data in enumerate(st.session_state.questions):
                n = i + 1
                q_text = q_data.get('question', '')
                ans_key = f"ans_{key}_{n}"

                # on_change 콜백으로 즉시 반영
                def make_callback(k, num):
                    def _cb():
                        val = st.session_state.get(f"ans_{k}_{num}", "")
                        st.session_state.answers[k][num] = val
                    return _cb

                # 현재 저장 상태 확인
                current_val = st.session_state.get(ans_key, st.session_state.answers[key][n])
                saved = bool(current_val.strip())

                col_q, col_badge = st.columns([5, 1])
                with col_q:
                    st.markdown(f'<div style="margin:8px 0 2px;font-weight:600">Q{n}.</div>', unsafe_allow_html=True)
                    st.code(q_text, language=None)
                with col_badge:
                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                    if saved:
                        st.markdown('<span class="mention-yes">✓ 저장됨</span>', unsafe_allow_html=True)
                    else:
                        st.markdown('<span class="mention-no">미입력</span>', unsafe_allow_html=True)

                st.text_area(
                    f"답변",
                    value=st.session_state.answers[key][n],
                    key=ans_key,
                    height=110,
                    placeholder=f"{label} 답변을 여기에 붙여넣으세요...",
                    label_visibility="collapsed",
                    on_change=make_callback(key, n)
                )
                st.markdown("<hr style='border:none;border-top:1px solid #f0f0f0;margin:4px 0'>", unsafe_allow_html=True)

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)
    col_back, col_next = st.columns([1, 3])
    with col_back:
        if st.button("← 질문 수정", use_container_width=True):
            st.session_state.step = 2
            st.rerun()
    with col_next:
        if st.button("📊 분석 시작 →", type="primary", use_container_width=True):
            if total == 0:
                st.error("최소 1개 이상의 답변을 입력해주세요.")
            else:
                # Claude API로 B2A 분석
                with st.spinner("Claude가 21개 답변을 분석 중입니다..."):
                    try:
                        client = get_client()
                        all_answers = []
                        for i, q_data in enumerate(st.session_state.questions):
                            n = i + 1
                            q_text = q_data.get('question', '')
                            for key, label in zip(ai_keys, ai_labels):
                                ans = st.session_state.answers[key][n]
                                if ans.strip():
                                    all_answers.append(f"[{label}] Q{n}: {q_text}\n답변: {ans}")

                        analysis_prompt = f"""당신은 GEO 분석 전문가입니다.
브랜드: {st.session_state.brand_name}
경쟁사: {st.session_state.brand_competitors}

아래는 ChatGPT(검색OFF/ON)와 Gemini에 실제 소비자 질문을 입력해서 얻은 AI 답변입니다.
각 답변을 분석해서 아래 형식으로 정리해주세요.

분석 항목:
1. 전체 요약: {st.session_state.brand_name}의 현재 AI 노출 현황 핵심 요약 (3~5문장)
2. 질문별 분석: Q1~Q7 각각에 대해
   - 브랜드 언급 여부 (O/X)
   - 언급 맥락 (긍정/중립/부정, 몇 번째 언급)
   - 경쟁사 언급 현황
   - 핵심 발견사항
3. 경쟁사 포지션: AI가 경쟁사를 어떻게 포지셔닝하는지
4. 핵심 공백: {st.session_state.brand_name}이 AI에서 완전히 빠진 영역
5. GEO 기회: 가장 빠르게 개선 가능한 포인트 3가지

--- 수집된 답변 ---
{chr(10).join(all_answers)}
"""
                        message = client.messages.create(
                            model="claude-sonnet-4-20250514",
                            max_tokens=3000,
                            messages=[{"role": "user", "content": analysis_prompt}]
                        )
                        st.session_state.analysis_result = message.content[0].text
                        st.session_state.cw_insights = [''] * 7
                        st.session_state.step = 4
                        st.rerun()

                    except Exception as e:
                        st.error(f"분석 오류: {str(e)}")

# ─────────────────────────────────────────────
# STEP 4: 분석 결과 + 인사이트
# ─────────────────────────────────────────────
elif st.session_state.step == 4:
    st.markdown('<div class="step-label">STEP 4</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-title">B2A 분석 결과 + 크림웍스 전략 인사이트</div>', unsafe_allow_html=True)

    # B2A 매트릭스
    st.markdown("#### 📊 B2A 매트릭스 — 언급 현황")
    ai_keys   = ['off', 'on', 'gem']
    ai_labels_short = ['GPT 검색OFF', 'GPT 검색ON', 'Gemini']

    header = st.columns([3, 1, 1, 1])
    header[0].markdown("**질문**")
    header[1].markdown("**GPT OFF**")
    header[2].markdown("**GPT ON**")
    header[3].markdown("**Gemini**")
    st.markdown("<hr style='border:none;border-top:1px solid #ddd;margin:4px 0'>", unsafe_allow_html=True)

    for i, q_data in enumerate(st.session_state.questions):
        n = i + 1
        q_text = q_data.get('question', '')
        row = st.columns([3, 1, 1, 1])
        row[0].markdown(f"**Q{n}.** {q_text[:38]}{'...' if len(q_text) > 38 else ''}")
        for j, key in enumerate(ai_keys):
            ans = st.session_state.answers[key][n]
            mention = check_mention(ans, st.session_state.brand_name)
            with row[j + 1]:
                if ans.strip():
                    if mention:
                        st.markdown('<span class="mention-yes">✓ 언급</span>', unsafe_allow_html=True)
                    else:
                        st.markdown('<span class="mention-no">✗ 미언급</span>', unsafe_allow_html=True)
                else:
                    st.caption("—")

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # Claude 분석 결과
    st.markdown("#### 🤖 Claude 분석 결과")
    st.markdown(f"""
    <div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:10px;
                padding:20px;font-size:0.88rem;line-height:1.8;white-space:pre-wrap;">
    {st.session_state.analysis_result}
    </div>""", unsafe_allow_html=True)

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # 크림웍스 전략 인사이트 편집
    st.markdown("#### 💜 크림웍스 전략 인사이트 편집")
    st.markdown('<div class="cw-box">보라색 영역은 크림웍스의 전략 의견입니다. 대표님이 직접 수정·보완해주세요.</div>', unsafe_allow_html=True)

    st.session_state.overall_diagnosis = st.text_area(
        "전체 현황 진단 요약",
        value=st.session_state.overall_diagnosis or st.session_state.analysis_result[:300],
        height=100,
        key="overall_diag"
    )

    for i, q_data in enumerate(st.session_state.questions):
        n = i + 1
        q_text = q_data.get('question', '')
        with st.expander(f"Q{n} 전략 인사이트 — {q_text[:40]}...", expanded=False):
            for key, label in zip(ai_keys, ai_labels_short):
                ans = st.session_state.answers[key][n]
                if ans.strip():
                    st.caption(f"**{label}:** {ans[:120]}{'...' if len(ans) > 120 else ''}")
            insight = st.text_area(
                "💜 크림웍스 전략 제안",
                value=st.session_state.cw_insights[i],
                key=f"cw_insight_{n}",
                height=90,
                placeholder="이 질문에 대한 GEO 전략 방향을 입력하세요..."
            )
            st.session_state.cw_insights[i] = insight

    st.session_state.priority_actions = st.text_area(
        "🚨 우선 실행 액션 플랜",
        value=st.session_state.priority_actions,
        height=140,
        key="priority_act"
    )

    col_back, col_next = st.columns([1, 3])
    with col_back:
        if st.button("← 답변 수정", use_container_width=True):
            st.session_state.step = 3
            st.rerun()
    with col_next:
        if st.button("📄 보고서 생성 →", type="primary", use_container_width=True):
            st.session_state.step = 5
            st.rerun()

# ─────────────────────────────────────────────
# STEP 5: Word 보고서 생성
# ─────────────────────────────────────────────
elif st.session_state.step == 5:
    st.markdown('<div class="step-label">STEP 5</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-title">Word 보고서 자동 생성</div>', unsafe_allow_html=True)

    def create_report():
        doc = Document()
        for section in doc.sections:
            section.top_margin    = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin   = Cm(2.5)
            section.right_margin  = Cm(2.5)

        brand_hex = st.session_state.brand_color
        br, bg, bb = hex_to_rgb(brand_hex)
        cr, cg, cb = CW_PURPLE_RGB
        ai_keys        = ['off', 'on', 'gem']
        ai_labels_s    = ['GPT 검색OFF', 'GPT 검색ON', 'Gemini']

        def add_run(para, text, bold=False, size=11, color=None, font="Arial"):
            run = para.add_run(text)
            run.bold = bold
            run.font.size = Pt(size)
            run.font.name = font
            if color:
                run.font.color.rgb = RGBColor(*color)
            return run

        # ── 표지 ──
        cover = doc.add_table(rows=1, cols=1)
        cover.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = cover.rows[0].cells[0]
        set_cell_bg(cell, '0f0f0f')
        for ptext, psize, pcolor, pbold, pspace in [
            ("CREAMWORKS  ×", 13, (br, bg, bb), True, (30, 8)),
            (st.session_state.brand_name, 32, (br, bg, bb), True, (0, 8)),
            ("AI 검색 최적화 (GEO) 전략 제안서", 17, (255, 255, 255), True, (0, 6)),
            (f"Presented by CREAMWORKS  ·  {datetime.now().strftime('%Y.%m')}", 10, (160, 160, 160), False, (20, 40)),
        ]:
            p = cell.add_paragraph() if ptext != "CREAMWORKS  ×" else cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(pspace[0])
            p.paragraph_format.space_after  = Pt(pspace[1])
            add_run(p, ptext, pbold, psize, pcolor)

        doc.add_page_break()

        # ── PART 0: 브랜드 프로파일 ──
        h = doc.add_heading("PART 0  브랜드 프로파일", level=1)
        h.runs[0].font.color.rgb = RGBColor(0, 0, 0)

        t = doc.add_table(rows=5, cols=2)
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        info_rows = [
            ("브랜드명", st.session_state.brand_name),
            ("카테고리", st.session_state.brand_category),
            ("핵심 USP", st.session_state.brand_usp),
            ("경쟁 브랜드", st.session_state.brand_competitors),
            ("강조 포인트", st.session_state.brand_focus),
        ]
        for row, (label, value) in zip(t.rows, info_rows):
            set_cell_bg(row.cells[0], brand_hex.replace('#', ''))
            p0 = row.cells[0].paragraphs[0]
            add_run(p0, label, bold=True, size=10, color=(0, 0, 0))
            p1 = row.cells[1].paragraphs[0]
            add_run(p1, value, size=10)
        doc.add_paragraph()

        # ── PART 1: 진단 질문 ──
        doc.add_page_break()
        h1 = doc.add_heading("PART 1  진단 질문 7개", level=1)
        h1.runs[0].font.color.rgb = RGBColor(0, 0, 0)

        for i, q_data in enumerate(st.session_state.questions):
            n = i + 1
            q_text = q_data.get('question', '')
            qtype  = q_data.get('type', '')
            check  = q_data.get('check_point', '')
            dlist  = q_data.get('data', [])

            qt = doc.add_table(rows=4 + len(dlist), cols=1)
            qt.style = 'Table Grid'

            # Q 헤더
            set_cell_bg(qt.rows[0].cells[0], '0f0f0f')
            add_run(qt.rows[0].cells[0].paragraphs[0], f"Q{n}.  {q_text}", bold=True, size=11, color=(br, bg, bb))

            # 유형
            set_cell_bg(qt.rows[1].cells[0], 'f5f5f5')
            add_run(qt.rows[1].cells[0].paragraphs[0], f"유형: {qtype}  |  확인 포인트: {check}", size=9, color=(80, 80, 80))

            # 데이터 헤더
            set_cell_bg(qt.rows[2].cells[0], '0f0f0f')
            add_run(qt.rows[2].cells[0].paragraphs[0], "📊  선정 근거 데이터", bold=True, size=10, color=(br, bg, bb))

            # 데이터 행
            for j, d in enumerate(dlist):
                set_cell_bg(qt.rows[3 + j].cells[0], 'ffffff' if j % 2 == 0 else 'fafafa')
                p = qt.rows[3 + j].cells[0].paragraphs[0]
                add_run(p, f"{d.get('source', '')} ({d.get('year', '')})  —  ", bold=True, size=9)
                add_run(p, d.get('content', ''), size=9)

            doc.add_paragraph()

        # ── PART 2: B2A 매트릭스 ──
        doc.add_page_break()
        h2 = doc.add_heading("PART 2  AI 진단 결과 — B2A 매트릭스", level=1)
        h2.runs[0].font.color.rgb = RGBColor(0, 0, 0)

        p_info = doc.add_paragraph()
        add_run(p_info, f"진단 기준: 3개 AI × 7개 질문 = 21개 답변  |  수집일: {datetime.now().strftime('%Y.%m.%d')}", size=9, color=(120, 120, 120))

        mt = doc.add_table(rows=8, cols=4)
        mt.style = 'Table Grid'
        for j, h_text in enumerate(["질문", "GPT 검색OFF", "GPT 검색ON", "Gemini"]):
            set_cell_bg(mt.rows[0].cells[j], '0f0f0f')
            add_run(mt.rows[0].cells[j].paragraphs[0], h_text, bold=True, size=10, color=(br, bg, bb))

        for i, q_data in enumerate(st.session_state.questions):
            n = i + 1
            row = mt.rows[n]
            add_run(row.cells[0].paragraphs[0], f"Q{n}. {q_data.get('question','')[:40]}", size=9)
            for j, key in enumerate(ai_keys):
                ans     = st.session_state.answers[key][n]
                mention = check_mention(ans, st.session_state.brand_name)
                cell    = row.cells[j + 1]
                if ans.strip():
                    if mention:
                        set_cell_bg(cell, 'd4edda')
                        add_run(cell.paragraphs[0], "✓  언급됨", size=9, color=(21, 87, 36))
                    else:
                        set_cell_bg(cell, 'f8d7da')
                        add_run(cell.paragraphs[0], "✗  미언급", size=9, color=(114, 28, 36))
                else:
                    add_run(cell.paragraphs[0], "—", size=9, color=(150, 150, 150))

        doc.add_paragraph()

        # ── PART 3: Claude 분석 결과 ──
        doc.add_page_break()
        h3 = doc.add_heading("PART 3  AI 진단 분석 결과", level=1)
        h3.runs[0].font.color.rgb = RGBColor(0, 0, 0)

        p_analysis = doc.add_paragraph()
        add_run(p_analysis, st.session_state.analysis_result, size=10)

        # ── PART 4: 크림웍스 전략 인사이트 ──
        doc.add_page_break()
        h4 = doc.add_heading("PART 4  크림웍스 GEO 전략 인사이트", level=1)
        h4.runs[0].font.color.rgb = RGBColor(cr, cg, cb)

        p_diag = doc.add_paragraph()
        add_run(p_diag, "전체 현황 진단", bold=True, size=11)
        doc.add_paragraph().add_run(st.session_state.overall_diagnosis).font.size = Pt(10)

        for i, q_data in enumerate(st.session_state.questions):
            n = i + 1
            insight = st.session_state.cw_insights[i]
            if insight.strip():
                it = doc.add_table(rows=2, cols=1)
                it.style = 'Table Grid'
                set_cell_bg(it.rows[0].cells[0], '0f0f0f')
                add_run(it.rows[0].cells[0].paragraphs[0],
                        f"Q{n}. {q_data.get('question', '')}", bold=True, size=10, color=(br, bg, bb))
                set_cell_bg(it.rows[1].cells[0], 'EDE9FF')
                add_run(it.rows[1].cells[0].paragraphs[0],
                        f"💜 크림웍스 전략: {insight}", size=10, color=(cr, cg, cb))
                doc.add_paragraph()

        # ── PART 5: 액션 플랜 ──
        doc.add_page_break()
        h5 = doc.add_heading("PART 5  우선 실행 액션 플랜", level=1)
        h5.runs[0].font.color.rgb = RGBColor(0, 0, 0)
        doc.add_paragraph().add_run(st.session_state.priority_actions).font.size = Pt(10)

        # 푸터
        fp = doc.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(fp, "CREAMWORKS  —  AI가 좋아하는 브랜드를 만듭니다", size=9, color=(150, 150, 150))

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    st.markdown("#### 보고서 구성")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        - **표지** — {st.session_state.brand_name} × CREAMWORKS
        - **PART 0** — 브랜드 프로파일
        - **PART 1** — 진단 질문 7개 + 근거 데이터
        """)
    with col2:
        st.markdown(f"""
        - **PART 2** — B2A 매트릭스 (21개 답변)
        - **PART 3** — Claude AI 분석 결과
        - **PART 4** — 크림웍스 전략 인사이트 (보라색)
        - **PART 5** — 우선 실행 액션 플랜
        """)

    st.markdown(f"""
    <div class="cw-box">
      💜 브랜드 컬러 <b>{st.session_state.brand_color}</b> + 크림웍스 퍼플 <b>#6B4EFF</b> 자동 적용
    </div>""", unsafe_allow_html=True)

    col_back, col_gen = st.columns([1, 3])
    with col_back:
        if st.button("← 인사이트 수정", use_container_width=True):
            st.session_state.step = 4
            st.rerun()
    with col_gen:
        if st.button("📄 Word 보고서 생성", type="primary", use_container_width=True):
            with st.spinner("보고서 생성 중..."):
                buf = create_report()
            fname = f"{st.session_state.brand_name}_GEO전략제안서_{datetime.now().strftime('%Y%m%d')}.docx"
            st.download_button(
                label=f"⬇️ {fname} 다운로드",
                data=buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            st.success("보고서가 생성되었습니다!")
            st.balloons()

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)
    if st.button("🔄 새 브랜드 진단 시작", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
