import streamlit as st
import anthropic
import io
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from docx.oxml.ns import qn
from docx.oxml import OxmlElement




# ─────────────────────────────────────────────
# Excel 불러오기 함수
# ─────────────────────────────────────────────
def load_from_excel(uploaded_file):
    """
    Excel 파일을 읽어 session_state에 복원.
    반환값: 'answer' (답변수집) 또는 'full' (전체데이터)
    """
    wb = openpyxl.load_workbook(uploaded_file)
    sheets = wb.sheetnames

    # ── 브랜드 정보 복원 ──
    if '브랜드 정보' in sheets:
        ws = wb['브랜드 정보']
        info_map = {
            '브랜드명':   'brand_name',
            '카테고리':   'brand_category',
            '핵심 USP':  'brand_usp',
            '주요 타겟':  'brand_target',
            '경쟁 브랜드': 'brand_competitors',
            '부정 이미지': 'brand_negative',
            '강조 포인트': 'brand_focus',
        }
        for row in ws.iter_rows(min_row=2, values_only=True):
            label, value = row[0], row[1]
            if label in info_map and value:
                st.session_state[info_map[label]] = str(value)

    # ── 질문 + 답변 복원 ──
    if 'AI 답변 수집' in sheets:
        ws2 = wb['AI 답변 수집']
        questions = []
        answers = {
            'off': {i: '' for i in range(1, 8)},
            'on':  {i: '' for i in range(1, 8)},
            'gem': {i: '' for i in range(1, 8)},
        }
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if not row[0] or not str(row[0]).startswith('Q'):
                continue
            n = int(str(row[0]).replace('Q', ''))
            # 컬럼: 번호(0), 질문(1), 유형(2), 단계(3), 확인포인트(4)
            #        GPT OFF(5), GPT ON(6), Gemini(7), 언급(8,9,10)
            q_dict = {
                'question':    str(row[1]) if row[1] else '',
                'type':        str(row[2]) if row[2] else '',
                'stage':       str(row[3]) if row[3] else '',
                'check_point': str(row[4]) if len(row) > 4 and row[4] else '',
                'data':        [],
            }
            questions.append(q_dict)
            # 새 형식 (확인포인트 컬럼 추가됨): 5,6,7
            # 구 형식 (확인포인트 없음): 4,5,6
            if len(row) >= 8 and row[5] is not None:
                answers['off'][n] = str(row[5]) if row[5] else ''
                answers['on'][n]  = str(row[6]) if row[6] else ''
                answers['gem'][n] = str(row[7]) if row[7] else ''
            else:
                answers['off'][n] = str(row[4]) if len(row) > 4 and row[4] else ''
                answers['on'][n]  = str(row[5]) if len(row) > 5 and row[5] else ''
                answers['gem'][n] = str(row[6]) if len(row) > 6 and row[6] else ''

        st.session_state.questions = questions
        st.session_state.answers   = answers
        st.session_state.questions_confirmed = True

    # ── Claude 분석 결과 복원 ──
    has_analysis = False
    if 'Claude 분석' in sheets:
        ws4 = wb['Claude 분석']
        for row in ws4.iter_rows(min_row=2, values_only=True):
            if row[0]:
                st.session_state.analysis_result = str(row[0])
                has_analysis = True
                break

    # ── 파일 타입 판별 ──
    if has_analysis and st.session_state.get('analysis_result', '').strip():
        # 전체데이터: STEP 4로
        st.session_state.overall_diagnosis  = st.session_state.analysis_result[:300]
        st.session_state.cw_insights        = [''] * 7
        st.session_state.priority_actions   = ''
        st.session_state.step = 4
        return 'full'
    else:
        # 답변수집: STEP 4로
        st.session_state.analysis_result    = ''
        st.session_state.overall_diagnosis  = ''
        st.session_state.cw_insights        = [''] * 7
        st.session_state.priority_actions   = ''
        st.session_state.step = 4
        return 'answer'



# ─────────────────────────────────────────────
# Excel 답변 저장 함수
# ─────────────────────────────────────────────
def create_answer_excel():
    br, bg, bb = brand_rgb()
    BRAND_HEX = st.session_state.brand_color.replace('#', 'FF')
    CW_HEX    = 'FFEADCF4'
    DARK_HEX  = 'FF1A1A1A'
    GRAY_HEX  = 'FFF7F7F7'
    GREEN_HEX = 'FFD9F2D0'
    RED_HEX   = 'FFFAE2D5'
    WHITE_HEX = 'FFFFFFFF'
    GRAY2_HEX = 'FFFAFAFA'
    # 헤더용: 브랜드 컬러 연하게 (흰 글씨 대신 어두운 글씨 사용)
    HEADER_HEX = f'FF{br:02X}{bg:02X}{bb:02X}'

    wb = openpyxl.Workbook()

    # ── 시트 1: 브랜드 정보 ──
    ws_info = wb.active
    ws_info.title = "브랜드 정보"

    header_fill = PatternFill("solid", fgColor=HEADER_HEX)
    brand_fill  = PatternFill("solid", fgColor=BRAND_HEX)
    cw_fill     = PatternFill("solid", fgColor=CW_HEX)
    gray_fill   = PatternFill("solid", fgColor=GRAY_HEX)
    white_fill  = PatternFill("solid", fgColor=WHITE_HEX)
    green_fill  = PatternFill("solid", fgColor=GREEN_HEX)
    red_fill    = PatternFill("solid", fgColor=RED_HEX)
    gray2_fill  = PatternFill("solid", fgColor=GRAY2_HEX)

    thin = Side(style='thin', color='DDDDDD')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def cell_style(ws, row, col, value, fill=None, font_color='FF000000',
                   bold=False, align='left', font_size=10):
        c = ws.cell(row=row, column=col, value=value)
        if fill:
            c.fill = fill
        c.font = Font(name='맑은 고딕', size=font_size,
                      bold=bold, color=font_color)
        c.alignment = Alignment(horizontal=align, vertical='center',
                                wrap_text=True)
        c.border = border
        return c

    # 브랜드 정보 헤더
    ws_info.merge_cells('A1:B1')
    c = ws_info['A1']
    c.value = f"{st.session_state.brand_name}  GEO 진단 — 브랜드 정보"
    c.fill = header_fill
    c.font = Font(name='맑은 고딕', size=13, bold=True, color='FFFFFFFF')
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws_info.row_dimensions[1].height = 30

    info_data = [
        ("브랜드명",    st.session_state.brand_name),
        ("카테고리",    st.session_state.brand_category),
        ("핵심 USP",   st.session_state.brand_usp),
        ("주요 타겟",  st.session_state.brand_target),
        ("경쟁 브랜드", st.session_state.brand_competitors),
        ("부정 이미지", st.session_state.brand_negative),
        ("강조 포인트", st.session_state.brand_focus),
        ("진단일",     datetime.now().strftime('%Y.%m.%d')),
    ]
    for i, (label, value) in enumerate(info_data):
        row = i + 2
        cell_style(ws_info, row, 1, label, brand_fill,
                   f'FF{br:02X}{bg:02X}{bb:02X}', True, 'left', 10)
        cell_style(ws_info, row, 2, value, gray_fill if i%2==0 else white_fill,
                   'FF333333', False, 'left', 10)
        ws_info.row_dimensions[row].height = 40

    ws_info.column_dimensions['A'].width = 18
    ws_info.column_dimensions['B'].width = 60

    # ── 시트 2: 질문 + 답변 Raw Data ──
    ws_ans = wb.create_sheet("AI 답변 수집")

    ai_keys   = ['off', 'on', 'gem']
    ai_labels = ['ChatGPT 검색OFF', 'ChatGPT 검색ON', 'Gemini']

    # 헤더 행
    headers = ["번호", "질문", "유형", "단계", "확인포인트",
               "ChatGPT 검색OFF", "ChatGPT 검색ON", "Gemini",
               "GPT OFF 언급", "GPT ON 언급", "Gemini 언급"]
    for ci, h in enumerate(headers):
        cell_style(ws_ans, 1, ci+1, h, header_fill,
                   'FFFFFFFF', True, 'center', 10)
    ws_ans.row_dimensions[1].height = 25

    # 데이터 행
    for i, q_data in enumerate(st.session_state.questions):
        n     = i + 1
        row   = i + 2
        bg_fill = gray_fill if i%2==0 else white_fill

        cell_style(ws_ans, row, 1, f"Q{n}", bg_fill, 'FF333333', True, 'center', 10)
        cell_style(ws_ans, row, 2, q_data.get('question',''), bg_fill, 'FF333333', False, 'left', 10)
        cell_style(ws_ans, row, 3, q_data.get('type',''), bg_fill, 'FF555555', False, 'left', 9)
        cell_style(ws_ans, row, 4, q_data.get('stage',''), bg_fill, 'FF555555', False, 'center', 9)
        cell_style(ws_ans, row, 5, q_data.get('check_point',''), bg_fill, 'FF555555', False, 'left', 9)

        for j, key in enumerate(ai_keys):
            ans = st.session_state.answers[key][n]
            cell_style(ws_ans, row, 6+j, ans, bg_fill, 'FF333333', False, 'left', 9)

            # 언급 여부
            mention = check_mention(ans, st.session_state.brand_name) if ans.strip() else None
            if mention is True:
                cell_style(ws_ans, row, 9+j, "✅ 언급됨", green_fill, 'FF155724', True, 'center', 9)
            elif mention is False:
                cell_style(ws_ans, row, 9+j, "❌ 미언급", red_fill, 'FFB43216', True, 'center', 9)
            else:
                cell_style(ws_ans, row, 9+j, "—", bg_fill, 'FF999999', False, 'center', 9)

        ws_ans.row_dimensions[row].height = 80

    col_widths = [8, 40, 25, 12, 35, 50, 50, 50, 12, 12, 12]
    for ci, w in enumerate(col_widths):
        ws_ans.column_dimensions[openpyxl.utils.get_column_letter(ci+1)].width = w

    # ── 시트 3: B2A 분석 결과 ──
    ws_b2a = wb.create_sheet("B2A 분석 결과")

    ws_b2a.merge_cells('A1:D1')
    c = ws_b2a['A1']
    c.value = f"B2A 매트릭스 — {st.session_state.brand_name}  |  수집일: {datetime.now().strftime('%Y.%m.%d')}"
    c.fill = header_fill
    c.font = Font(name='맑은 고딕', size=12, bold=True, color='FFFFFFFF')
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws_b2a.row_dimensions[1].height = 28

    b2a_headers = ["질문", "GPT 검색OFF", "GPT 검색ON", "Gemini"]
    for ci, h in enumerate(b2a_headers):
        cell_style(ws_b2a, 2, ci+1, h, cw_fill, 'FF5327A8', True, 'center', 10)
    ws_b2a.row_dimensions[2].height = 22

    for i, q_data in enumerate(st.session_state.questions):
        n   = i + 1
        row = i + 3
        bg_fill = gray2_fill if i%2==0 else white_fill

        cell_style(ws_b2a, row, 1,
                   f"Q{n}. {q_data.get('question','')}",
                   bg_fill, 'FF333333', False, 'left', 9)

        for j, key in enumerate(ai_keys):
            ans     = st.session_state.answers[key][n]
            mention = check_mention(ans, st.session_state.brand_name) if ans.strip() else None
            if mention is True:
                cell_style(ws_b2a, row, 2+j, "✅ 언급됨", green_fill, 'FF155724', True, 'center', 9)
            elif mention is False:
                cell_style(ws_b2a, row, 2+j, "❌ 미언급", red_fill, 'FFB43216', True, 'center', 9)
            else:
                cell_style(ws_b2a, row, 2+j, "—", bg_fill, 'FF999999', False, 'center', 9)

        ws_b2a.row_dimensions[row].height = 30

    ws_b2a.column_dimensions['A'].width = 50
    for col in ['B','C','D']:
        ws_b2a.column_dimensions[col].width = 18

    # ── 시트 4: Claude 분석 결과 ──
    ws_cl = wb.create_sheet("Claude 분석")
    ws_cl.merge_cells('A1:B1')
    c = ws_cl['A1']
    c.value = "Claude B2A 분석 결과"
    c.fill = header_fill
    c.font = Font(name='맑은 고딕', size=12, bold=True, color='FFFFFFFF')
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws_cl.row_dimensions[1].height = 28

    analysis_text = st.session_state.get('analysis_result', '')
    c2 = ws_cl.cell(row=2, column=1, value=analysis_text)
    c2.font = Font(name='맑은 고딕', size=9)
    c2.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    c2.fill = PatternFill("solid", fgColor=CW_HEX)
    ws_cl.row_dimensions[2].height = 300
    ws_cl.column_dimensions['A'].width = 100

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf





import json
from datetime import datetime

# ─────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="GEO-Scan v2 | 크림웍스",
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

  /* 연두색 액션 버튼 */
  div[data-testid="stButton"] button[kind="primary"] {{
      background-color: #52B788 !important;
      border-color: #52B788 !important;
      color: white !important;
      font-weight: 600 !important;
      max-width: 320px !important;
      border-radius: 8px !important;
  }}
  div[data-testid="stButton"] button[kind="primary"]:hover {{
      background-color: #40916C !important;
      border-color: #40916C !important;
  }}
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
        'brand_color': '#4A90D9',
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
  <div style="font-size:0.78rem;color:#888;letter-spacing:2px;font-weight:600;">CREAMWORKS  ·  GEO-Scan  |  앱 1</div>
  <div style="font-size:1.7rem;font-weight:800;margin:6px 0 4px;">{bn} AI 진단 시스템</div>
  <div style="font-size:0.85rem;color:#bbb;">브랜드 AI 검색 노출 현황 진단 · 분석 결과 Excel 저장</div>
</div>
""", unsafe_allow_html=True)

# 스텝 인디케이터
step_names = ["브랜드 입력", "질문 확인", "답변 수집", "분석·Excel 저장"]
cols = st.columns(4)
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
# 이전 작업 이어하기 (Excel 업로드)
# ─────────────────────────────────────────────
if st.session_state.step == 1:
    with st.expander("📂 이전 작업 이어하기 (Excel 업로드)", expanded=False):
        st.caption("이전에 저장한 Excel 파일을 업로드하면 중간부터 이어서 진행할 수 있습니다.")

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**케이스 A — 답변수집 Excel**")
            st.caption("브랜드 정보 + 질문 + 답변 복원 → STEP 4 (분석) 부터 시작")
        with col_b:
            st.markdown("**케이스 B — 전체데이터 Excel**")
            st.caption("브랜드 정보 + 질문 + 답변 + 분석 복원 → STEP 4 (분석·Excel 저장) 부터 시작")

        uploaded_xl = st.file_uploader(
            "Excel 파일 업로드 (답변수집 또는 전체데이터)",
            type=["xlsx"],
            key="resume_upload"
        )

        if uploaded_xl is not None:
            try:
                file_type = load_from_excel(uploaded_xl)
                if file_type == 'answer':
                    st.success(f"✅ **답변수집 파일** 복원 완료! → STEP 4 (분석 시작) 로 이동합니다.")
                    brand = st.session_state.get('brand_name','')
                    st.info(f"브랜드: **{brand}** | 질문 **{len(st.session_state.questions)}개** | 답변 복원 완료")
                else:
                    st.success(f"✅ **전체데이터 파일** 복원 완료! → STEP 4 (분석·Excel 저장) 로 이동합니다.")
                    brand = st.session_state.get('brand_name','')
                    st.info(f"브랜드: **{brand}** | Claude 분석 결과 포함")

                if st.button("▶ 이어서 진행하기", type="primary"):
                    st.rerun()

            except Exception as e:
                st.error(f"파일 읽기 오류: {e}\n올바른 GEO-Scan Excel 파일을 업로드해주세요.")

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
        <div style="font-size:0.78rem;color:#aaa;margin-bottom:4px">
          구글에 "{st.session_state.brand_name or '브랜드명'} 브랜드 컬러 HEX" 검색 후 입력하세요
        </div>
        <div style="font-size:0.78rem;color:#aaa;margin-bottom:8px">
          예) 교촌치킨 → #F9BA15 &nbsp;|&nbsp; 드시모네 → #01C49C &nbsp;|&nbsp; 크림웍스 → #7C5CBF
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

    col_btn, col_empty = st.columns([2, 3])
    with col_btn:
        clicked = st.button("🔍  질문 7개 생성하기", type="primary", use_container_width=True)

    if clicked:
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

                    brand_nm    = st.session_state.brand_name
                    brand_comp  = st.session_state.brand_competitors
                    brand_cat   = st.session_state.brand_category

                    prompt = f"""당신은 GEO(Generative Engine Optimization) 전문가입니다.
아래 브랜드 정보를 바탕으로, 소비자가 ChatGPT·Gemini에 실제로 물어볼 법한 GEO 진단 질문 7개를 설계해주세요.

⚠️ 핵심 주의: 각 필드에는 질문 설계 메타데이터만 작성합니다. 절대로 AI 답변 내용, 추천 목록, 제품 설명을 작성하지 마세요.

브랜드 정보:
- 브랜드명: {brand_nm}
- 카테고리: {st.session_state.brand_category}
- 핵심 USP: {st.session_state.brand_usp}
- 주요 타겟: {st.session_state.brand_target}
- 경쟁 브랜드: {brand_comp}
- 부정 이미지/약점: {st.session_state.brand_negative}
- 강조 포인트: {st.session_state.brand_focus}

질문 설계 원칙:
1. 브랜드명이 절대 들어가면 안 됨 (브랜드를 모르는 소비자가 AI에게 묻는 질문)
2. 실제 소비자가 쓰는 구어체
3. 단계: DISCOVER(2개) / CONSIDER(3개) / DECIDE(2개)
4. 브랜드 USP와 직접 연결되는 질문 우선
5. 부정 이미지 방어 질문 1개 이상 포함

[check_point 필드 작성 규칙 — 반드시 준수]
- 반드시 1~2문장의 짧은 진단 기준만 작성
- 형식 예시: "{brand_nm}이(가) [구체적 맥락]으로 언급되는지 + {brand_comp} 대비 포지션 언급 여부"
- 절대 금지: AI 답변 내용, 제품 설명, 마크다운(#, *, -), URL, 추천 목록 등 일절 포함 금지
- 최대 길이: 100자 이내

[data 필드 작성 규칙]
- 각 질문마다 실제 기관/매체 데이터 3개
- "출처기관" 같은 템플릿 문자열 절대 금지. 반드시 실제 기관명 사용
- 실제 기관 예시: 한국소비자원, 식품의약품안전처, 오픈서베이, 네이버 데이터랩, 닐슨코리아, 통계청, aT한국농수산식품유통공사, 와이즈앱 등
- content는 구체적 수치/통계 포함, 1~2문장 이내
- year는 2023~2026 사이 실제 연도

반드시 아래 JSON 형식으로만 응답 (마크다운 코드블록 없이, JSON만):
{{
  "questions": [
    {{
      "question": "질문 내용 (구어체, 브랜드명 없이)",
      "stage": "DISCOVER 또는 CONSIDER 또는 DECIDE",
      "type": "유형명 10자 이내",
      "check_point": "1~2문장 진단 기준만. 마크다운/답변내용/URL 절대 금지. 100자 이내.",
      "data": [
        {{"source": "실제기관명", "content": "구체적 수치 포함 1~2문장", "year": "2024"}},
        {{"source": "실제기관명", "content": "구체적 수치 포함 1~2문장", "year": "2025"}},
        {{"source": "실제기관명", "content": "구체적 수치 포함 1~2문장", "year": "2024"}}
      ]
    }}
  ]
}}"""

                    message = client.messages.create(
                        model="claude-sonnet-4-20250514",
                        max_tokens=8000,
                        messages=[{"role": "user", "content": prompt}]
                    )

                    raw = message.content[0].text.strip()

                    # JSON 추출 - 코드블록 및 앞뒤 텍스트 제거
                    import re as _re
                    cb = _re.search(r'```(?:json)?\s*(\{.*?\})\s*```', raw, _re.DOTALL)
                    if cb:
                        raw = cb.group(1)
                    else:
                        m = _re.search(r'(\{.*\})', raw, _re.DOTALL)
                        if m:
                            raw = m.group(1)
                    raw = raw.strip()

                    parsed = json.loads(raw)
                    questions = parsed['questions']

                    # check_point 오염 방지: 200자 초과 or 마크다운 감지 시 자동 정제
                    for q in questions:
                        cp = q.get('check_point', '')
                        if len(cp) > 200 or any(c in cp for c in ['#', '**', '[', 'http', '---', '\n\n']):
                            first = cp.split('\n')[0][:150].strip()
                            q['check_point'] = first if first else f"{st.session_state.brand_name}의 AI 노출 현황 확인"

                    st.session_state.questions = questions
                    st.session_state.questions_confirmed = False
                    st.session_state.step = 2
                    st.rerun()

                except json.JSONDecodeError as je:
                    st.error(f"질문 생성 중 JSON 파싱 오류가 발생했습니다. 다시 시도해주세요.\n{str(je)}")
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

    col_back, col_confirm, col_empty2 = st.columns([1, 2, 2])
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

    # Excel 저장 버튼
    st.markdown("#### 💾 답변 저장")
    st.caption("수집한 답변을 중간 저장해두세요. 세션 종료 시 데이터가 사라집니다. 저장한 파일은 다시 업로드해서 이어서 진행할 수 있습니다.")

    if st.button("📊 답변 Excel 저장", use_container_width=True):
        if total == 0:
            st.error("저장할 답변이 없습니다. 먼저 답변을 입력해주세요.")
        else:
            with st.spinner("Excel 파일 생성 중..."):
                buf = create_answer_excel()
            fname = f"{st.session_state.brand_name}_GEO_답변수집_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                label=f"⬇️ {fname} 다운로드",
                data=buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_excel_ans"
            )
            st.success(f"Excel 파일이 생성되었습니다! ({total}/21개 답변 저장)")

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)
    col_back, col_next, col_empty3 = st.columns([1, 2, 2])
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

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # ── Excel 최종 저장 ──
    st.markdown("#### 💾 GEO-Report 앱용 Excel 저장")
    st.markdown('<div class="cw-box">💜 분석이 완료됐습니다. Excel을 저장한 뒤 <b>GEO-Report 앱</b>에 업로드하면 Word·PPT 보고서가 자동 생성됩니다.</div>', unsafe_allow_html=True)

    if st.button("📊 전체 데이터 Excel 저장", type="primary", use_container_width=True):
        with st.spinner("Excel 파일 생성 중..."):
            buf = create_answer_excel()
        fname = f"{st.session_state.brand_name}_GEO_전체데이터_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label=f"⬇️ {fname} 다운로드",
            data=buf,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_excel_full"
        )
        st.success("✅ Excel 저장 완료! GEO-Report 앱에서 열어주세요.")

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)
    col_back, col_new = st.columns([1, 3])
    with col_back:
        if st.button("← 답변 수정", use_container_width=True):
            st.session_state.step = 3
            st.rerun()
    with col_new:
        if st.button("🔄 새 브랜드 진단 시작", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
