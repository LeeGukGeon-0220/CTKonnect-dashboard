# -*- coding: utf-8 -*-
"""
CTKonnect 실적 대시보드 (캐나다 대시보드와 같은 모양 · 2026-10-29)
--------------------------------------------------------------
  * 캐나다 app.py 를 바탕으로 CTKonnect, Inc.(미국) 용으로 바꾼 것입니다.
  * 기준 자료 = 연결패키지(26년 원장_raw · BS_IS_매핑 · BS · IS · 조직운영안 · 기타주요사항 시트).
  * 통화는 USD 하나 · SR&ED · 재고수불부(원가 기준) 없음.
--------------------------------------------------------------
v7 (2026-08-20) 전면 개편
  * 모든 숫자를 '원장(raw)'에서 직접 계산합니다. 엑셀 수식에 의존하지 않습니다.
  * 거래일 DD/MM/YYYY 텍스트를 명시적으로 파싱합니다 (일/월 뒤바뀜 방지).
  * 매출원가 / 판관비 / 영업이익 계산 오류를 바로잡았습니다.
  * '계정과목 상세' 페이지를 새로 만들었습니다.
--------------------------------------------------------------
"""
import io
import os
import re
import math
import zipfile
import posixpath
import collections
import json
import html as _html
import calendar
import datetime
import urllib.request

import pandas as pd
import numpy as np
import altair as alt
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title='CTKonnect 실적보고', layout='wide')
os.chdir(os.path.dirname(os.path.abspath(__file__)))   # 어디서 켜도 이 폴더의 자료를 읽습니다

# ══════════════════════════════════════════════════════════════
# 00. 들어갈 때 비밀번호 (OTC 대시보드와 같은 방식 · 2026-10-26)
#     ※ 회사 재무자료라 아무나 못 보게 막습니다. 비밀번호는 아래 한 줄에서 바꿉니다.
#       (비밀번호는 코드에 두지 않고 Streamlit 설정 Secrets 의 dashboard_password 를 씁니다)
# ══════════════════════════════════════════════════════════════
대시보드비번 = ''
try:
    대시보드비번 = str(st.secrets.get('dashboard_password', '') or '').strip()
except Exception:
    pass
if not 대시보드비번:
    st.error('비밀번호 설정이 없습니다. Streamlit 설정(Secrets) 또는 .streamlit/secrets.toml 에 dashboard_password 를 넣어 주세요.')
    st.stop()
if not st.session_state.get('들어옴'):
    st.html("""<style>
[data-testid="stSidebar"], [data-testid="collapsedControl"] {display:none}
.login-box {max-width:420px; margin:14vh auto 0; padding:34px 36px 26px; border-radius:14px 14px 0 0;
            background:#12294D; color:#FFFFFF; font-family:'Malgun Gothic','맑은 고딕',sans-serif;
            box-shadow:0 8px 30px rgba(13,27,51,0.25)}
.login-box .t {font-size:21px; font-weight:800; letter-spacing:-0.3px}
.login-box .s {font-size:13px; color:rgba(255,255,255,0.72); margin-top:6px; line-height:1.7}
.login-box .bar {height:3px; width:46px; background:#E2611C; border-radius:2px; margin:14px 0 4px}
</style>
<div class="login-box"><div class="t">CTKonnect 실적보고</div>
<div class="s">회계팀 내부 자료입니다 — 비밀번호를 넣고 들어와 주세요.</div>
<div class="bar"></div></div>""")
    _l, 가운데, _r = st.columns([1, 1.1, 1])
    with 가운데:
        with st.form('login', clear_on_submit=False, border=False):
            넣은것 = st.text_input('비밀번호', type='password', label_visibility='collapsed',
                                   placeholder='비밀번호')
            누름 = st.form_submit_button('들어가기', type='primary', width='stretch')
        if 누름:
            if 넣은것.strip() == 대시보드비번:
                st.session_state['들어옴'] = True
                st.rerun()
            else:
                st.error('비밀번호가 다릅니다.')
    st.stop()

# ══════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════
# 0. 테마 / 스타일  (밝은 화면 · 어두운 화면 두 가지)
# ══════════════════════════════════════════════════════════════
if 'ui_theme' not in st.session_state:
    st.session_state['ui_theme'] = '밝은 화면'
밝게 = st.session_state['ui_theme'] == '밝은 화면'

# ── 색 기준 ────────────────────────────────────────────────────
#  남색(navy)   : 화면의 뼈대 — 사이드바 · 표 머리글 · 제목
#  주황(accent) : 시선을 끌 곳 하나만 — 현재 메뉴 · 제목 밑줄 · 당해연도 막대
#  숫자는 검정, 음수는 (빨강 괄호) — 이 규칙은 두 화면에서 같습니다.
if 밝게:
    T = dict(bg='#F6F7F9', panel='#FFFFFF', panel2='#FAFBFD', line='#E7E7E1',
             ink='#0B0B0B', ink2='#4A4A46', ink3='#8A8880',
             accent='#E2611C', accentbg='rgba(226,97,28,0.10)', dot='#C7C6BC',
             grid='rgba(11,11,11,0.07)', axis='#8A8880',
             ok='#0A6B3D', warn='#D03B3B', shadow='0 1px 2px rgba(13,27,51,0.05)',
             head='#12294D', zebra='#FAFBFD',
             side='#12294D', upc='#D03B3B', downc='#1D3E70',
             navy='#12294D', navy7='#1D3E70', navysoft='#F1F4F9', navyline='#D8E0EC',
             headink='#0D1B33', strongline='#12294D',
             sideink='#FFFFFF', sideink2='rgba(255,255,255,0.76)',
             sideink3='rgba(255,255,255,0.50)', sideline='rgba(255,255,255,0.13)',
             sidehover='rgba(255,255,255,0.09)')
else:
    T = dict(bg='#0B1220', panel='#111C30', panel2='#0E1829', line='rgba(255,255,255,0.10)',
             ink='#F2F5FA', ink2='#B9C4D6', ink3='#8798AE',
             accent='#F0834F', accentbg='rgba(240,131,79,0.16)', dot='#3C4E69',
             grid='rgba(255,255,255,0.07)', axis='#8798AE',
             ok='#3DD68C', warn='#F06A6A', shadow='0 1px 3px rgba(0,0,0,0.35)',
             head='#1D3E70', zebra='rgba(255,255,255,0.035)',
             side='#0A1526', upc='#F06A6A', downc='#7FA8E0',
             navy='#16305C', navy7='#2B5490', navysoft='rgba(255,255,255,0.05)',
             navyline='rgba(255,255,255,0.14)',
             headink='#F2F5FA', strongline='rgba(255,255,255,0.28)',
             sideink='#F2F5FA', sideink2='rgba(255,255,255,0.74)',
             sideink3='rgba(255,255,255,0.48)', sideline='rgba(255,255,255,0.11)',
             sidehover='rgba(255,255,255,0.08)')

# 그래프 계열 색 — 1·2·3번은 색약(色弱)에서도 서로 구분되도록 검증된 조합입니다.
#   TEAL  = 당해연도 · 대표 계열 (주황)      SLATE = 전년 · 비교용 (회색)
#   BLUE  = 두 번째 계열 (파랑)              AMBER = 세 번째 계열 (청록)
#   ROSE  = 경고 · 마이너스 (빨강)
if 밝게:
    TEAL, ROSE, BLUE, SLATE, AMBER = '#eb6834', '#D03B3B', '#2a78d6', '#898781', '#1baf7a'
else:
    TEAL, ROSE, BLUE, SLATE, AMBER = '#d95926', '#F06A6A', '#3987e5', '#8E8C84', '#199e70'


def _ctk_theme():
    return {
      'padding': {'left': 46, 'top': 6, 'right': 18, 'bottom': 6},
      'config': {
        'background': 'transparent',
        'view': {'stroke': 'transparent'},
        'font': "'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo',sans-serif",
        'axis': {'domainColor': T['navyline'], 'gridColor': T['grid'],
                 'labelColor': T['ink3'], 'titleColor': T['ink2'], 'tickColor': T['navyline'],
                 'labelFontSize': 13, 'titleFontSize': 13, 'labelPadding': 5,
                 'labelFontWeight': 500, 'grid': True, 'domain': True},
        'legend': {'labelColor': T['ink2'], 'titleColor': T['ink2'], 'labelFontWeight': 600,
                   'labelFontSize': 13, 'titleFontSize': 13, 'symbolSize': 118,
                   'symbolType': 'square'},
        'text': {'color': T['headink'], 'fontWeight': 700, 'fontSize': 12.5},
        'title': {'color': T['headink'], 'fontSize': 16.5, 'fontWeight': 800, 'anchor': 'start',
                  'offset': 12},
    }}


alt.themes.register('ctk_theme', _ctk_theme)
alt.themes.enable('ctk_theme')

st.markdown(f"""
<style>
/* ── 글꼴 : 화면 전체를 맑은 고딕으로 통일 ────────────────── */
html, body,
[data-testid="stAppViewContainer"] *,
[data-testid="stSidebar"] * {{
    font-family: 'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo',
                 system-ui, -apple-system, sans-serif !important; }}
/* ★아이콘은 글꼴이 곧 그림이라 반드시 되돌려 놓습니다 — 안 그러면 화살표가
   'keyboard_double_arrow_left' 같은 글자로 보입니다. (위 규칙보다 뒤에 둡니다) */
[data-testid="stIconMaterial"], [data-testid="stIconMaterial"] *,
[data-testid="stTooltipIcon"], [data-testid="stTooltipIcon"] *,
.stTooltipIcon, .stTooltipIcon * {{
    font-family: 'Material Symbols Rounded' !important; }}
[data-testid="stAppViewContainer"], [data-testid="stHeader"] {{ background-color: {T['bg']}; }}
.main .block-container, [data-testid="stMainBlockContainer"], .block-container {{
    color: {T['ink']};
    padding-top: 1.6rem !important; padding-bottom: 2.5rem !important;
    padding-left: 0.9rem !important; padding-right: 1.2rem !important;
    max-width: none !important; }}
/* 오른쪽 위 Streamlit 메뉴(점 세 개 ⋮)는 다시 보이게 둡니다 — 공유·인쇄·설정이 거기 있습니다.
   맨 위 무지개 줄만 감추고, 머리줄은 배경 없이 띄워 화면은 그대로 깔끔하게 둡니다. */
[data-testid="stDecoration"] {{ display: none !important; }}
[data-testid="stHeader"] {{ background: transparent !important; }}
[data-testid="stToolbar"] {{ display: flex !important; right: 0.6rem !important; }}
[data-testid="stToolbar"] button {{ color: {T['ink3']} !important; }}
/* 요소 사이 세로 간격을 조금 좁힙니다 */
.main [data-testid="stVerticalBlock"] {{ gap: 0.7rem !important; }}
[data-testid="stSidebar"] {{ background-color: {T['side']};
    border-right: 1px solid {T['line']}; }}
[data-testid="stSidebar"] * {{ color: {T['sideink']} !important; }}
.brand-box {{ padding: 10px 4px 18px 4px; border-bottom: 1px solid {T['sideline']}; margin-bottom: 10px; }}
.brand-title {{ font-size: 17.5px; font-weight: 800; letter-spacing: -0.03em;
    color: {T['sideink']} !important; }}
.brand-sub {{ font-size: 11px; font-weight: 700; color: {T['sideink3']} !important;
    letter-spacing: 0.12em; }}
/* 사이드바 제목 아래 「최종자료 2026년 8월」 — 경영진이 언제까지 자료인지 바로 알도록 */
.asof-badge {{ margin: -2px 2px 12px 2px; padding: 8px 11px; border-radius: 9px;
    background: {T['accentbg']}; border: 1px solid {T['accent']};
    font-size: 12.5px; font-weight: 700; color: {T['sideink']} !important;
    letter-spacing: -0.01em; line-height: 1.5; }}
.asof-badge b {{ color: {T['accent']} !important; font-size: 14px; }}
.asof-badge .sub {{ display: block; font-size: 11px; font-weight: 600;
    color: {T['sideink3']} !important; letter-spacing: 0.02em; margin-top: 2px; }}
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button {{
    background: transparent !important; border: none !important; border-radius: 9px !important;
    box-shadow: none !important; padding: 10px 12px !important; min-height: 0 !important;
    display: flex !important; align-items: center !important; }}
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button > div {{
    justify-content: flex-start !important; width: 100% !important; }}
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button,
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button * {{
    color: {T['sideink2']} !important; font-size: 14.5px !important; font-weight: 600 !important;
    text-align: left !important; }}
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button::before {{
    content: ""; display: inline-block; width: 6px; height: 6px; border-radius: 50%;
    background: rgba(255,255,255,0.34); margin-right: 11px; flex-shrink: 0; }}
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button:hover {{
    background: {T['sidehover']} !important; }}
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button:hover,
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button:hover * {{
    color: {T['sideink']} !important; }}
/* 지금 보고 있는 메뉴 — 왼쪽에 주황 막대를 세워 색을 못 알아봐도 구분되게 합니다 */
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button[kind="primary"] {{
    background: {T['sidehover']} !important;
    border-left: 3px solid {T['accent']} !important; }}
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button[kind="primary"],
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button[kind="primary"] * {{
    color: {T['sideink']} !important; font-weight: 800 !important; }}
[data-testid="stSidebar"] .st-key-nav_container [data-testid="stButton"] > button[kind="primary"]::before {{
    background: {T['accent']} !important; }}
[data-testid="stSidebar"] .st-key-pl_sub [data-testid="stButton"] > button,
[data-testid="stSidebar"] .st-key-rep_sub [data-testid="stButton"] > button {{
    padding-left: 30px !important; }}
[data-testid="stSidebar"] .st-key-pl_sub [data-testid="stButton"] > button::before,
[data-testid="stSidebar"] .st-key-rep_sub [data-testid="stButton"] > button::before {{
    content: "-"; width: auto; height: auto; border-radius: 0;
    background: transparent !important; color: {T['ink3']};
    font-weight: 700; margin-right: 9px; line-height: 1; }}
[data-testid="stSidebar"] .st-key-pl_sub [data-testid="stButton"]
    > button[kind="primary"]::before,
[data-testid="stSidebar"] .st-key-rep_sub [data-testid="stButton"]
    > button[kind="primary"]::before {{
    color: {T['accent']} !important; background: transparent !important;
    width: auto !important; height: auto !important; border-radius: 0 !important; }}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
    font-size: 12.5px !important; color: {T['sideink3']} !important; }}
[data-testid="stSidebar"] hr {{ border-color: {T['sideline']} !important;
    margin: 14px 0 10px !important; }}
.side-foot {{ margin-top: 6px; padding-top: 13px; border-top: 1px solid {T['sideline']};
    font-size: 12.5px; line-height: 1.9; color: {T['sideink3']}; }}
.side-foot b {{ color: {T['sideink']}; font-weight: 700; }}
/* 설정 접이칸 — 평소엔 조용히 있다가 누르면 펴집니다 */
[data-testid="stSidebar"] [data-testid="stExpander"] {{
    border: 1px solid {T['sideline']} !important; border-radius: 9px !important;
    background: transparent !important; }}
[data-testid="stSidebar"] [data-testid="stExpander"] summary {{
    padding: 8px 12px !important; }}
/* 펼쳤을 때 제목줄이 흰 바탕이 되어 글씨가 안 보이던 것을 막습니다 */
[data-testid="stSidebar"] [data-testid="stExpander"] details,
[data-testid="stSidebar"] [data-testid="stExpander"] summary,
[data-testid="stSidebar"] [data-testid="stExpander"] summary:hover,
[data-testid="stSidebar"] [data-testid="stExpanderDetails"] {{
    background: transparent !important; background-color: transparent !important; }}
[data-testid="stSidebar"] [data-testid="stExpander"] summary svg {{
    fill: {T['sideink2']} !important; color: {T['sideink2']} !important; }}
[data-testid="stSidebar"] [data-testid="stExpander"] summary p {{
    font-size: 12.5px !important; font-weight: 700 !important;
    color: {T['sideink2']} !important; }}
/* 사이드바가 남색이라 안에 들어가는 입력칸도 어둡게 맞춥니다 (안 그러면 흰 글씨가 안 보입니다) */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{
    background: {T['sidehover']} !important; border: 1px dashed {T['sideink3']} !important; }}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {{
    color: {T['sideink2']} !important; }}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {{
    background: {T['accent']} !important; border: none !important; }}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button,
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * {{
    color: #FFFFFF !important; font-weight: 700 !important; }}
[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] {{
    background: {T['sidehover']} !important; border-radius: 8px; padding: 6px 8px; }}
[data-testid="stSidebar"] input, [data-testid="stSidebar"] [data-baseweb="input"],
[data-testid="stSidebar"] [data-baseweb="base-input"],
[data-testid="stSidebar"] [data-testid="stNumberInputContainer"],
[data-testid="stSidebar"] [data-baseweb="select"] > div {{
    background: {T['sidehover']} !important; color: {T['sideink']} !important;
    border-color: {T['sideline']} !important; }}
[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"],
[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"] {{
    background: {T['sidehover']} !important; }}
[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"] svg,
[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"] svg {{
    fill: {T['sideink2']} !important; }}
[data-testid="stSidebar"] [data-baseweb="radio"] div[aria-checked="false"] {{
    border-color: {T['sideink3']} !important; background: transparent !important; }}
h1, h2, h3, h4 {{ color: {T['headink']} !important; font-weight: 800 !important;
    letter-spacing: -0.03em !important; }}
.main .block-container p, .main .block-container li,
.main .block-container [data-testid="stMarkdownContainer"] {{ font-size: 16px; }}
[data-testid="stCaptionContainer"] p {{ font-size: 14.5px !important; color: {T['ink3']} !important; }}
.chart-head {{ display:flex; align-items:baseline; justify-content:space-between;
    margin: 6px 0 2px; }}
.chart-head .t {{ font-size:17.5px; font-weight:800; letter-spacing:-.025em; color:{T['headink']};
    padding-bottom:7px; position:relative; }}
.chart-head .t::after {{ content:''; position:absolute; left:0; bottom:0; width:26px; height:3px;
    background:{T['accent']}; border-radius:2px; }}
.chart-head .u {{ font-size:14px; font-weight:600; color:{T['ink3']}; }}
/* 어두운 화면에서 선택상자·라디오가 혼자 하얗게 남지 않도록 맞춥니다 */
{'' if 밝게 else f'''
.main [data-baseweb="select"] > div, .main [data-testid="stNumberInputContainer"],
.main [data-baseweb="popover"] li, .main [data-baseweb="menu"] {{
    background: {T['panel']} !important; color: {T['ink']} !important;
    border-color: {T['line']} !important; }}
.main [data-baseweb="select"] svg {{ fill: {T['ink2']} !important; }}
.main [data-testid="stRadio"] label p {{ color: {T['ink2']} !important; }}
/* 본문 파일 업로드 칸 (실적보고 엑셀작성 화면) 도 함께 어둡게 */
[data-testid="stMainBlockContainer"] [data-testid="stFileUploaderDropzone"] {{
    background: {T['panel2']} !important; border: 1px dashed {T['line']} !important; }}
[data-testid="stMainBlockContainer"] [data-testid="stFileUploaderDropzone"] * {{
    color: {T['ink2']} !important; }}
[data-testid="stMainBlockContainer"] [data-testid="stFileUploaderDropzone"] button {{
    background: {T['accent']} !important; border: none !important; }}
[data-testid="stMainBlockContainer"] [data-testid="stFileUploaderDropzone"] button,
[data-testid="stMainBlockContainer"] [data-testid="stFileUploaderDropzone"] button * {{
    color: #FFFFFF !important; font-weight: 700 !important; }}
[data-testid="stMainBlockContainer"] [data-testid="stFileUploaderFile"] {{
    background: {T['panel2']} !important; border-radius: 8px; padding: 6px 8px; }}
[data-testid="stMainBlockContainer"] [data-testid="stFileUploaderFile"] * {{
    color: {T['ink']} !important; }}
'''}
</style>
""", unsafe_allow_html=True)

카드스타일 = f"""
<style>
  .wrap {{ margin:0 0 6px; padding:2px; color:{T['ink']};
          font-family:'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo',
                      system-ui,-apple-system,sans-serif; }}
  .page-head {{ display:flex; align-items:center; gap:12px; margin:0 0 3px; }}
  .page-head .t {{ font-size:26px; font-weight:800; letter-spacing:-.035em;
      color:{T['headink']}; line-height:1.2; }}
  .page-head .pill {{ background:{T['navy']}; color:#FFFFFF; font-size:13px; font-weight:800;
      letter-spacing:-.01em; padding:6px 15px; border-radius:22px; white-space:nowrap; }}
  .page-cap {{ font-size:14px; color:{T['ink3']}; margin:0 0 16px; }}
  .wrap .note {{ font-size:13.5px; color:{T['ink3']}; font-weight:600; margin-bottom:12px; }}
  .wrap .kpi-row {{ display:grid; grid-template-columns:repeat(var(--n,4),1fr); gap:12px; margin-bottom:16px; }}
  .wrap .kpi-card {{ background:{T['panel']}; border:1px solid {T['line']}; border-radius:12px;
              padding:0 17px 15px; box-shadow:{T['shadow']}; overflow:hidden; }}
  /* 카드 위에 얇은 띠 — 첫 칸만 주황으로 눈에 띄게 합니다 */
  .wrap .kpi-card::before {{ content:''; display:block; height:3px; margin:0 -17px 13px;
              background:{T['navyline']}; }}
  .wrap .kpi-row .kpi-card:first-child::before {{ background:{T['accent']}; }}
  .wrap .kpi-label {{ font-size:13.5px; font-weight:700; color:{T['headink']}; margin-bottom:10px;
              letter-spacing:-.01em; }}
  .wrap .kpi-value {{ font-size:34px; font-weight:800; color:{T['ink']};
               letter-spacing:-.035em; line-height:1.05; }}
  .wrap .kpi-value.ok {{ color:{T['ok']}; }} .kpi-value.warn {{ color:{T['warn']}; }}
  .wrap .kpi-delta {{ margin-top:9px; font-size:14px; }}
  /* 실적보고 맨 위 손익 흐름 줄 — 작게 */
  .wrap .kpi-card.mini {{ padding:0 10px 7px; border-radius:9px; box-shadow:none; }}
  .wrap .kpi-card.mini::before {{ height:2px; margin:0 -10px 6px; }}
  .wrap .kpi-card.mini .kpi-label {{ font-size:11.5px; margin-bottom:3px; }}
  .wrap .kpi-card.mini .kpi-value {{ font-size:17px !important; font-weight:800; }}
  .wrap .kpi-card.mini .kpi-value .unit {{ font-size:10.5px; }}
  .wrap .kpi-card.mini .kpi-delta {{ margin-top:3px; font-size:10.5px !important; line-height:1.4; }}
  .wrap .unit {{ font-size:13.5px; font-weight:600; color:{T['ink3']}; }}
  .wrap .up {{ color:{T['upc']}; }} .down {{ color:{T['downc']}; }} .muted {{ color:{T['ink3']}; }}
  .wrap .card {{ background:{T['panel']}; border:1px solid {T['line']}; border-radius:12px;
          padding:15px 20px 16px; margin-bottom:12px; box-shadow:{T['shadow']}; }}
  /* 제목 아래 짧은 주황 밑줄 — 어디를 보는 화면인지 바로 알 수 있게 */
  .wrap .card h3 {{ font-size:18px; font-weight:800; color:{T['headink']}; margin:0 0 3px;
          letter-spacing:-.025em; padding-bottom:7px; position:relative; }}
  .wrap .card h3::after {{ content:''; position:absolute; left:0; bottom:0; width:26px; height:3px;
          background:{T['accent']}; border-radius:2px; }}
  .wrap .card .sub {{ font-size:13.5px; color:{T['ink3']}; margin:8px 0 11px; }}
  .wrap table {{ width:100%; border-collapse:separate; border-spacing:0; }}
  /* 남색 머리칸 글자는 모두 가운데 정렬합니다 (표·세부내역 공통) */
  .wrap th {{ text-align:center; font-size:13.5px; color:#FFFFFF; background:{T['navy']};
       font-weight:800; padding:9px 11px; letter-spacing:-.01em; white-space:nowrap; }}
  .wrap thead tr:first-child th:first-child {{ border-radius:8px 0 0 0; }}
  .wrap thead tr:first-child th:last-child {{ border-radius:0 8px 0 0; }}
  .wrap td:first-child {{ text-align:left; }}
  .wrap td {{ padding:8px 11px; font-size:15px; line-height:1.3; border-bottom:1px solid {T['line']};
       text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap;
       color:{T['ink2']}; }}
  .wrap td:first-child {{ color:{T['ink']}; font-weight:500; }}
  .wrap tbody tr:last-child td {{ border-bottom:none; }}
  .wrap tbody tr:last-child td:first-child {{ border-radius:0 0 0 8px; }}
  .wrap tbody tr:last-child td:last-child {{ border-radius:0 0 8px 0; }}
  .wrap td.lft {{ text-align:left; }}   /* 머리칸(th)은 가운데, 내용칸만 왼쪽 */
  .wrap tbody tr:hover td {{ background:{T['zebra']}; }}
  .wrap tr.total td {{ font-weight:800; background:{T['zebra']}; color:{T['ink']}; }}
  .wrap tr.sub td:first-child {{ padding-left:22px; color:{T['ink2']}; font-weight:400; }}
  .wrap tr.sub2 td:first-child {{ padding-left:44px; color:{T['ink3']}; }}
  .wrap .tag {{ display:inline-block; font-size:12.5px; padding:1px 8px; border-radius:20px; margin-left:8px; font-weight:700; }}
  .wrap .tag.rev {{ color:{T['accent']}; background:{T['accentbg']}; }}
  .wrap .tag.cogs {{ color:{AMBER}; background:rgba(217,138,23,0.15); }}
  .wrap .tag.sga {{ color:{BLUE}; background:rgba(59,130,246,0.15); }}
  .wrap .tag.non {{ color:#8B7BE8; background:rgba(139,123,232,0.15); }}
  .wrap .bars {{ display:flex; align-items:flex-end; gap:18px; height:178px; padding:0 4px; }}
  .wrap .bar-group {{ display:flex; flex-direction:column; align-items:center; gap:8px; flex:1; height:100%; justify-content:flex-end; }}
  .wrap .bar-pair {{ display:flex; gap:7px; align-items:flex-end; height:152px; }}
  .wrap .bar-col {{ display:flex; flex-direction:column; align-items:center; justify-content:flex-end; }}
  .wrap .bar-val {{ font-size:13px; font-weight:600; color:{T['ink2']}; margin-bottom:4px;
             font-variant-numeric:tabular-nums; white-space:nowrap; }}
  .wrap .bar {{ width:17px; border-radius:4px 4px 0 0; }}
  .wrap .bar.cur {{ background:{TEAL}; }} .bar.prev {{ background:{SLATE}; }}
  .wrap .bar-month {{ font-size:14px; color:{T['ink3']}; }}
  .wrap .bar-sep {{ width:1px; align-self:stretch; margin:0 6px 26px;
             background:{T['line']}; flex:0 0 1px; }}
  .wrap .bar-group.cum .bar-month {{ font-weight:700; color:{T['ink']}; }}
  .wrap .bar-group.cum .bar-val {{ color:{T['ink']}; }}
  .wrap .unitbadge {{ float:right; font-size:13px; font-weight:600; color:{T['ink3']}; }}
  .wrap .mix-total {{ display:flex; justify-content:space-between; font-size:15.5px; font-weight:800;
               color:{T['headink']}; border-top:2px solid {T['strongline']};
               padding-top:11px; margin-top:14px; }}
  /* ── 셋트로 나뉜 표 (당월·월별 실적집계) ─────────────────── */
  .wrap table.fixed {{ width:auto; table-layout:fixed; }}
  .wrap table.fixed th, .wrap table.fixed td {{ width:128px; }}
  /* 머리글은 <br> 로만 줄을 나누고, 위아래 가운데에 놓습니다 */
  .wrap table.fixed thead th {{ white-space:nowrap; line-height:1.5;
      vertical-align:middle; padding-top:7px; padding-bottom:7px; }}
  .wrap table.fixed {{ border:none; }}
  .wrap table.fixed th {{ text-align:center; }}
  .wrap table.fixed td.name {{ text-align:left; }}   /* 머리칸은 가운데 */
  .wrap table.fixed th.name, .wrap table.fixed td.name {{ width:210px; }}
  .wrap table.fixed th.grp {{ padding-bottom:6px; }}
  /* 셋트 경계선 */
  .wrap table.fixed th.s1, .wrap table.fixed td.s1 {{ border-left:1px solid {T['navyline']}; }}
  .wrap table.fixed th.gh {{ text-align:center; font-size:13px; font-weight:800;
      color:#FFFFFF; background:{T['navy7']}; padding:7px 9px; letter-spacing:-.01em; }}
  .wrap table.fixed th.gh0 {{ background:{T['navy']}; vertical-align:middle; }}
  /* 증감 칸만 주황 — 강조색은 한 곳에만 씁니다 */
  .wrap table.fixed th.d {{ background:{T['accent']}; color:#FFFFFF; }}
  .wrap table.fixed td.d {{ background:{T['accentbg']}; font-weight:700; }}
  .wrap table.fixed td.d.plus {{ color:{T['upc']}; }}
  .wrap table.fixed td.d.minus {{ color:{T['downc']}; }}
  .wrap table.fixed td.d.zero {{ color:{T['ink3']}; font-weight:400; }}
  /* 표가 화면보다 넓으면 옆으로 밀어 볼 수 있게 하고, 스크롤 막대를 눈에 보이게 둡니다 */
  .wrap .scrollx {{ overflow-x:auto; scrollbar-width:thin; padding-bottom:3px; }}
  .wrap .scrollx::-webkit-scrollbar {{ height:9px; }}
  .wrap .scrollx::-webkit-scrollbar-track {{ background:transparent; }}
  .wrap .scrollx::-webkit-scrollbar-thumb {{ background:{T['line']}; border-radius:5px; }}
  .wrap .scrollx:hover::-webkit-scrollbar-thumb {{ background:{T['ink3']}; }}
  /* 긴 표 — 아래로 내려도 남색 머리줄(계정·연도)이 늘 맨 위에 붙어 있게 합니다 */
  .wrap .stick {{ overflow:auto; max-height:74vh; scrollbar-width:thin; }}
  .wrap .stick::-webkit-scrollbar {{ width:9px; height:9px; }}
  .wrap .stick::-webkit-scrollbar-track {{ background:transparent; }}
  .wrap .stick::-webkit-scrollbar-thumb {{ background:{T['line']}; border-radius:5px; }}
  .wrap .stick:hover::-webkit-scrollbar-thumb {{ background:{T['ink3']}; }}
  .wrap .stick thead th {{ position:sticky; top:0; z-index:4; line-height:1.35;
      background:{T['navy']}; box-shadow:inset 0 -1px 0 {T['navy']}; }}
  .wrap .stick thead th .th2 {{ font-size:11.5px; font-weight:600; opacity:0.85; }}
  .wrap .stick thead th:last-child {{ background:{T['navy7']}; }}
  /* 현금흐름 — 앞으로 올 달(예상) 칸은 색을 달리해 실적과 구분합니다 */
  .wrap table.lined th.fc {{ background:{T['navy7']}; line-height:1.25; }}
  .wrap table.lined th.fc .fcs {{ font-size:11px; font-weight:600; opacity:0.85; }}
  .wrap table.lined td.fc {{ background:{T['navysoft']}; }}
  /* CTKonnect — 금액이 커서(수천 K USD) 월 칸 숫자가 넘치지 않게 글자 · 여백을 조금 줄입니다 */
  .wrap table.lined th, .wrap table.lined td {{ font-size:13.5px; padding-left:4px !important; padding-right:4px !important;
      font-variant-numeric:tabular-nums; letter-spacing:-.02em; }}
  /* 국문 ↔ 영문 대사 알림 */
  .wrap .reconc {{ margin:10px 0 2px; padding:10px 13px; border-radius:8px;
      background:{T['navysoft']}; font-size:13px; line-height:1.7; color:{T['ink2']}; }}
  .wrap .reconc b {{ color:{T['headink']}; font-weight:800; }}
  .wrap .chks {{ display:flex; flex-wrap:wrap; gap:6px; margin-top:6px; }}
  .wrap .chk {{ display:inline-block; padding:3px 9px; border-radius:999px;
      font-size:12.5px; font-weight:700; }}
  .wrap .chk.ok {{ background:{T['panel2']}; color:{T['ink']}; }}
  .wrap .chk.no {{ background:{T['warn']}; color:#FFFFFF; }}
  .wrap tr.head td {{ font-weight:700; color:{T['headink']}; }}
  .wrap .neg {{ color:{T['warn']}; }}
  .wrap .memo {{ display:flex; justify-content:space-between; gap:10px;
      font-size:12.5px; line-height:1.6; padding:5px 0;
      border-bottom:1px dashed {T['line']}; }}
  .wrap .memo b {{ color:{T['ink']}; font-weight:600; flex:0 0 auto; max-width:52%;
      overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
  .wrap .memo span {{ color:{T['ink3']}; text-align:right; }}
  /* 거래 내용 — 거래처 × 달 표. 옆(달)·아래(거래처) 어느 쪽으로도 밀어 볼 수 있고,
     맨 윗줄과 맨 왼쪽 칸은 밀어도 제자리에 붙어 있습니다. */
  .wrap .mscroll {{ overflow:auto; max-height:330px; scrollbar-width:thin;
      border:1px solid {T['line']}; border-radius:8px; }}
  .wrap .mscroll::-webkit-scrollbar {{ width:9px; height:9px; }}
  .wrap .mscroll::-webkit-scrollbar-track {{ background:transparent; }}
  .wrap .mscroll::-webkit-scrollbar-thumb {{ background:{T['line']}; border-radius:5px; }}
  .wrap .mscroll:hover::-webkit-scrollbar-thumb {{ background:{T['ink3']}; }}
  .wrap table.mtab {{ border-collapse:separate; border-spacing:0; font-size:12.5px;
      width:100%; }}
  .wrap table.mtab th {{ position:sticky; top:0; z-index:3; background:{T['navy']};
      color:#FFFFFF; font-weight:700; padding:7px 7px; text-align:center;
      white-space:nowrap; }}
  .wrap table.mtab td {{ padding:6px 7px; text-align:right; white-space:nowrap;
      color:{T['ink2']}; border-top:1px solid {T['line']}; }}
  .wrap table.mtab th.lft, .wrap table.mtab td.lft {{ position:sticky; left:0;
      text-align:left; width:124px; min-width:124px; max-width:124px;
      white-space:normal; }}
  .wrap table.mtab th.lft {{ z-index:5; }}
  .wrap table.mtab td.lft {{ z-index:2; background:{T['panel']};
      box-shadow:1px 0 0 {T['line']}; }}
  .wrap table.mtab td.lft b {{ color:{T['ink']}; font-weight:600; font-size:11.5px;
      display:block; line-height:1.35; }}
  .wrap table.mtab td.lft .cap {{ display:block; font-size:11px; font-weight:400;
      color:{T['ink3']}; line-height:1.4; white-space:nowrap;
      overflow:hidden; text-overflow:ellipsis; }}
  .wrap table.mtab td.sum, .wrap table.mtab th.sum {{ font-weight:700;
      color:{T['ink']}; background:{T['navysoft']}; }}
  .wrap table.mtab th.sum {{ color:#FFFFFF; background:{T['navy7']}; }}
  .wrap table.mtab tr.tot td {{ font-weight:700; color:{T['ink']};
      background:{T['panel2']}; border-top:2px solid {T['line']}; }}
  .wrap table.mtab tr.tot td.lft {{ background:{T['panel2']}; }}
  .wrap .calc {{ font-size:13px; font-weight:400; color:{T['ink3']}; letter-spacing:0; }}
  .wrap .calcnote {{ margin-top:13px; padding:12px 14px; border-radius:8px;
      background:{T['navysoft']}; border:none;
      font-size:13.5px; line-height:1.75; color:{T['ink2']}; }}
  .wrap .calcnote b {{ color:{T['headink']}; font-weight:800; }}
  .wrap .calcnote ul.fnote {{ margin:4px 0 0; padding-left:18px; }}
  .wrap .calcnote ul.fnote li {{ margin:1px 0; }}
  .wrap .calcnote li.red {{ color:{T['warn']}; font-weight:700; }}
  /* 예상(추정) 칸이 어떤 가정으로 나온 값인지 — 실적이 아니라 가정임을 붉게 못박습니다 */
  .wrap .assume {{ margin-top:13px; padding:12px 14px; border-radius:8px;
      background:{T['accentbg']}; border-left:4px solid {T['warn']};
      font-size:13.5px; line-height:1.8; color:{T['ink2']}; }}
  .wrap .assume b {{ color:{T['headink']}; font-weight:800; }}
  .wrap .assume .red {{ color:{T['warn']}; font-weight:800; }}
  .wrap .assume .num {{ color:{T['warn']}; font-weight:700;
      font-variant-numeric:tabular-nums; }}
  .wrap .assume ul {{ margin:6px 0 0; padding-left:19px; }}
  .wrap .assume li {{ margin:2px 0; }}
  .wrap tr.total td .calc {{ font-weight:400; }}
  /* 판관비 세부 줄 — 「판관비 합계」 를 누르면 펼쳐집니다 (숨은 체크박스) */
  .wrap input.togbox {{ display:none; }}
  .wrap table.lined tr.sga {{ display:none; }}
  .wrap input.togbox:checked ~ .scrollx table.lined tr.sga {{ display:table-row; }}
  .wrap label.tog {{ cursor:pointer; }}
  .wrap label.tog .caret {{ display:inline-block; color:{T['accent']}; margin-left:4px; transition:transform .15s; }}
  .wrap input.togbox:checked ~ .scrollx label.tog .caret {{ transform:rotate(90deg); }}
  /* 주석 · 지난 달 — 누르면 펼쳐짐 */
  .wrap details.fold > summary {{ cursor:pointer; list-style:none; }}
  .wrap details.fold > summary::-webkit-details-marker {{ display:none; }}
  .wrap details.fold > summary::before {{ content:'▸ '; color:{T['accent']}; }}
  .wrap details.fold[open] > summary::before {{ content:'▾ '; }}
  .wrap .eff {{ margin-top:13px; padding:12px 14px; border-radius:8px; border:1.5px solid {T['accent']};
      background:{T['accentbg']}; font-size:13.5px; line-height:1.7; color:{T['ink2']}; }}
  .wrap .eff h4 {{ margin:0 0 4px; font-size:15px; color:{T['accent']}; font-weight:900; }}
  .wrap .eff table {{ width:100%; table-layout:fixed; margin-top:6px; background:#fff; }}
  .wrap .eff td {{ vertical-align:top; }}
  /* 공헌이익 — 가장 눈에 띄게 (주황 테두리 · 연한 주황 바탕 · 큰 글씨) */
  .wrap table.lined tr.hl td {{ background:{T['accentbg']} !important; font-weight:900; font-size:15.5px;
      border-top:2px solid {T['accent']}; border-bottom:2px solid {T['accent']}; }}
  .wrap table.lined tr.hl td:first-child {{ color:{T['accent']}; }}
  .wrap table.lined tr.hl td .calc {{ font-weight:600; color:{T['ink2']}; font-size:12.5px; }}
  /* 손익계산서 · 현금흐름 — 화면이 좁아지면 열도 같이 줄어듭니다 (열너비는 %로 잡습니다).
     두 표가 같은 %를 쓰기 때문에 화면 크기가 달라져도 위아래 칸이 항상 맞습니다. */
  .wrap table.lined {{ width:100%; table-layout:fixed; }}
  .wrap table.lined th, .wrap table.lined td {{ padding-left:7px; padding-right:7px; }}
  .wrap table.lined th:first-child, .wrap table.lined td:first-child {{ white-space:normal; line-height:1.4; }}
  .wrap table.lined th {{ text-align:center; }}
  /* 맨 오른쪽 「누적」 칸은 제일 자주 보는 곳이라 따로 세웁니다 */
  .wrap table.lined th:last-child {{ background:{T['navy7']}; }}
  .wrap table.lined td:last-child {{ background:{T['navysoft']}; color:{T['ink']}; font-weight:800; }}
  .wrap table.lined tr.total td:last-child {{ background:{T['navysoft']};
      filter:brightness(0.97); }}
  /* 오른쪽 세부내역 — 거래처 이름이 길어도 카드 밖으로 삐져나가지 않게 */
  .wrap .card table {{ max-width:100%; }}
  /* 현금흐름 카드 안의 월말 잔액 꺾은선 — 아래 표의 월 칸과 가로 위치를 맞춥니다.
     선은 viewBox 를 100 으로 두고 preserveAspectRatio="none" 으로 늘려 %와 정확히 맞추고,
     굵기가 같이 늘어나지 않도록 vector-effect="non-scaling-stroke" 를 씁니다. */
  .wrap .cashline {{ position:relative; width:100%; height:126px; margin:2px 0 12px; }}
  .wrap .cashline .cgrid {{ position:absolute; height:0; border-top:1px solid {T['grid']}; }}
  .wrap .cashline .cgrid.zero {{ border-top-color:{T['navyline']}; }}
  /* 세로 눈금 글자는 그림 맨 왼쪽에서 축선 바로 앞까지 오른쪽 맞춤으로 앉힙니다
     (기초 잔액 점이 축선 위에 서므로 겹치지 않게 왼쪽으로 밀어 둡니다) */
  .wrap .cashline .cylab {{ position:absolute; left:0; transform:translateY(-50%);
      font-size:12.5px; font-weight:500;
      color:{T['ink3']}; text-align:right; font-variant-numeric:tabular-nums; }}
  .wrap .cashline .cdomain {{ position:absolute; top:0; bottom:0; width:0;
      border-left:1px solid {T['navyline']}; }}
  /* ※ Streamlit 이 <svg> 를 지워 버려서, 선과 면을 칸(div) + clip-path 로 그립니다.
       칸 하나가 점과 점 사이 한 구간이고, 네 꼭짓점을 %로 찍어 사다리꼴을 오려 냅니다. */
  .wrap .cashline .cseg, .wrap .cashline .cfill {{ position:absolute; top:0; height:100%; }}
  .wrap .cashline .cseg {{ background:{T['accent']}; }}
  .wrap .cashline .cdot {{ position:absolute; width:9px; height:9px; border-radius:50%;
      background:{T['accent']}; transform:translate(-50%,-50%);
      box-shadow:0 0 0 2px {T['panel']}; }}
  .wrap .cashline .cdot.last {{ width:11px; height:11px; }}
  /* 기초 잔액 점은 속을 비워, 「출발점이지 이번 달 실적이 아니다」를 눈으로 구분합니다 */
  .wrap .cashline .cdot.first {{ background:{T['panel']};
      box-shadow:0 0 0 2px {T['accent']}; }}
  .wrap .cashline .cval {{ position:absolute; transform:translate(-50%,-100%);
      font-size:10px; font-weight:800; color:{T['ink']}; white-space:nowrap;
      font-variant-numeric:tabular-nums; }}
  .wrap .cashline .ctag {{ position:absolute; transform:translate(-50%,0);
      margin-top:9px; font-size:11.5px; font-weight:700; color:{T['ink3']};
      white-space:nowrap; }}
  .wrap .cashline .cbase {{ position:absolute; height:0;
      border-top:1px solid {T['navyline']}; }}
  /* 그림 제목은 그림 바로 위 한 줄로 둡니다 (왼쪽 자리는 기초 점과 눈금 글자가 씁니다) */
  .wrap .cashcap {{ margin:4px 0 2px; display:flex; align-items:baseline; gap:10px;
      flex-wrap:wrap; }}
  .wrap .cashcap .t {{ font-size:15px; font-weight:800;
      color:{T['headink']}; letter-spacing:-.02em; line-height:1.35; }}
  .wrap .cashcap .u {{ font-size:12.5px; font-weight:600; color:{T['ink3']}; }}
  .wrap table.detail {{ table-layout:fixed; }}
  .wrap table.detail th, .wrap table.detail td {{ padding-left:5px; padding-right:5px; }}
  .wrap td.lft {{ white-space:normal; word-break:break-word; }}
  /* 화면이 아주 좁을 때만 글자를 줄입니다.
     ※ 예전에는 1500px 이하에서 바로 줄여 버려, 보통 크기 창에서도 글자가 작게 나왔습니다.
        기준을 1280px 로 낮춰 웬만한 화면에서는 제 크기(15px)가 그대로 보이게 했습니다. */
  @media (max-width:1280px) {{
    .wrap table.lined th, .wrap table.lined td {{ font-size:14px; padding-left:6px; padding-right:6px; }}
    .wrap .card table {{ font-size:14px; }}
    .wrap table.fixed th, .wrap table.fixed td {{ width:106px; font-size:14px; padding-left:6px; padding-right:6px; }}
    .wrap table.fixed th.name, .wrap table.fixed td.name {{ width:178px; }}
  }}
  @media (max-width:1080px) {{
    .wrap table.lined th, .wrap table.lined td {{ font-size:13px; padding-left:5px; padding-right:5px; }}
    .wrap table.fixed th, .wrap table.fixed td {{ width:94px; font-size:13px; padding-left:5px; padding-right:5px; }}
    .wrap table.fixed th.name, .wrap table.fixed td.name {{ width:162px; }}
  }}
  .wrap .legend {{ display:flex; gap:16px; margin-top:12px; font-size:14px; color:{T['ink2']}; }}
  .wrap .legend span {{ display:flex; align-items:center; gap:5px; }}
  .wrap .sw {{ width:10px; height:10px; border-radius:2px; display:inline-block; }}
  .wrap .mix-item .top {{ display:flex; justify-content:space-between; font-size:15.5px; margin-bottom:5px; }}
  .wrap .mix-track {{ height:8px; background:{T['grid']}; border-radius:5px; overflow:hidden; margin-bottom:12px; }}
  .wrap .mix-fill {{ height:100%; background:{TEAL}; border-radius:5px; }}
  .wrap .warnbox {{ background:{T['accentbg']}; border-left:4px solid {T['accent']};
             border-radius:9px; padding:13px 15px; font-size:14px; color:{T['ink2']};
             font-weight:600; margin-bottom:14px; }}
  /* 데이터 점검 — 한 줄에 아이콘 · 항목 · 값 */
  .wrap .chk {{ display:flex; align-items:center; gap:12px; padding:12px 14px; border-radius:9px;
             margin-bottom:8px; border:1px solid {T['line']}; background:{T['panel2']}; }}
  .wrap .chk.ok {{ border-left:4px solid {T['ok']}; }}
  .wrap .chk.wn {{ border-left:4px solid {T['accent']}; background:{T['accentbg']}; }}
  .wrap .chk .ic {{ width:21px; height:21px; flex:0 0 21px; border-radius:50%; color:#FFFFFF;
             display:flex; align-items:center; justify-content:center;
             font-size:12px; font-weight:800; }}
  .wrap .chk.ok .ic {{ background:{T['ok']}; }}
  .wrap .chk.wn .ic {{ background:{T['accent']}; }}
  .wrap .chk .t {{ flex:1; font-size:14.5px; font-weight:600; color:{T['ink2']}; }}
  .wrap .chk .v {{ font-size:14.5px; font-weight:800; color:{T['ink']};
             font-variant-numeric:tabular-nums; }}
  /* 정상 대 장기미수 비율 띠 */
  .wrap .split {{ display:flex; height:40px; border-radius:8px; overflow:hidden;
             gap:2px; margin:2px 0 16px; }}
  .wrap .split i {{ display:flex; align-items:center; justify-content:center; color:#FFFFFF;
             font-size:12.5px; font-weight:800; font-style:normal; min-width:6px; }}
  .wrap .split i.ok {{ background:{T['navy7']}; }}
  .wrap .split i.bad {{ background:{ROSE}; }}
  .wrap td .sw {{ width:10px; height:10px; border-radius:3px; display:inline-block;
             margin-right:8px; vertical-align:middle; }}
</style>
"""

st.markdown(카드스타일, unsafe_allow_html=True)   # 카드 스타일은 여기서 한 번만
CARD_CSS = ''  # (아래 f-string 들과의 호환을 위해 빈 값으로 둡니다)

매출계정 = ['제품매출', '기타매출', '용역수익', '상품매출', '용역매출']
매출원가계정 = ['제품매출원가', '상품매출원가', '용역매출원가', '패키지매출원가', '기타매출원가']   # 2026-11-03 : 씨티케이커넥트 용역매출원가
영업외계정 = ['이자비용', '이자수익', '잡손실', '잡이익', '외환차손익', '외화환산손익', '세금환급', '수수료수익']
판관비항목 = ['인건비성 항목', '커미션 비용', '외부용역', '건물관리비', '차량유지비', 'ICT',
             '출장비', '운반 및 창고료', '소모성경비', '기타관리경비']   # CTKonnect : 상각비 없음 (2026-10-29)
ACT2ITEM = {'인건비성 항목': '인건비성 항목', '커미션 비용': '커미션 비용', '외부용역': '외부용역',
            '건물': '건물관리비', '차량': '차량유지비', 'ICT': 'ICT', '출장비': '출장비',
            '운반 및 창고료': '운반 및 창고료', '소모성경비': '소모성경비'}


# ══════════════════════════════════════════════════════════════
# 1. 데이터 읽기 (v7 신규 양식 / v6 이전 양식 자동 판별)
# ══════════════════════════════════════════════════════════════
def _parse_dmy(v):
    """QuickBooks 원장의 '일/월/연도' 텍스트를 진짜 날짜로 바꿉니다."""
    # ※ 빈 칸은 float(nan) 으로 올 때도 있고 NaT 으로 올 때도 있습니다.
    #   NaT 은 datetime 의 한 종류라서 아래 isinstance 에 걸려 버리므로 여기서 먼저 걸러 냅니다.
    if v is None:
        return pd.NaT
    try:
        if pd.isna(v):
            return pd.NaT
    except (TypeError, ValueError):
        pass
    if isinstance(v, (datetime.datetime, pd.Timestamp)):
        return pd.Timestamp(v).normalize()
    if isinstance(v, datetime.date):
        return pd.Timestamp(v)
    parts = str(v).strip().replace('-', '/').split('/')
    if len(parts) != 3:
        return pd.NaT
    try:
        return pd.Timestamp(int(parts[2]), int(parts[1]), int(parts[0]))
    except Exception:
        return pd.NaT


def _표준화(df, 매핑표, 연도힌트):
    """v6 / v7 어느 양식이든 같은 컬럼 이름으로 맞춥니다."""
    ren = {'Distribution account': '계정영문', 'Transaction date': '거래일',
           'Transaction ID': 'TransactionID', '#': '전표번호',
           '계정분류(BS/IS)': '계정분류'}
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    if '계정영문' not in df.columns:
        return None
    df = df[df['계정영문'].notna()].copy()
    df['계정영문'] = df['계정영문'].astype(str).str.strip()
    df = df[~df['계정영문'].isin(['', 'nan', 'Beginning Balance'])]

    df['거래일'] = df['거래일'].apply(_parse_dmy)
    df = df[df['거래일'].notna()].copy()

    for c in ('Debit', 'Credit'):
        df[c] = pd.to_numeric(df.get(c), errors='coerce').fillna(0)
    df['금액'] = df['Credit'] - df['Debit']          # 수익 +, 비용 −
    df['년'] = df['거래일'].dt.year
    df['월'] = df['거래일'].dt.month

    # 계정분류/계정과목은 항상 매핑표 기준으로 다시 계산 (원본 수식 오류 방지)
    k = df['계정영문'].str.lower()
    df['계정분류'] = k.map(매핑표['bsis'])
    df['계정과목'] = k.map(매핑표['kor'])

    def _분류(a):
        if pd.isna(a):
            return np.nan
        if a in 매출계정:
            return '매출'
        if a in 매출원가계정:
            return '매출원가'
        if a in 영업외계정:
            return '영업외손익'
        if a == '법인세비용':          # CTKonnect : 세전이익 아래 — 판관비 · 영업외손익에 넣지 않음
            return '법인세'
        return '판관비'

    df['분류'] = np.where(df['계정분류'].eq('IS'), df['계정과목'].map(_분류), None)

    # 원장 양식에 따라 열 이름이 「활동분류」이기도 하고 「활동분류(대분류)」이기도 합니다.
    # ★둘 중 어느 쪽이 와도 같게 읽어야 합니다★ — 안 그러면 판관비가 전부 「기타관리경비」로 잡힙니다.
    if '활동분류(대분류)' not in df.columns and '활동분류' in df.columns:
        df['활동분류(대분류)'] = df['활동분류']
    for c in ('활동분류(대분류)', '활동분류(소분류)', '활동세부', '분류상태', 'Description', 'Name', 'Balance'):
        if c not in df.columns:
            df[c] = np.nan

    def _보고항목(r):
        if r['분류'] != '판관비':
            return np.nan
        if r['계정과목'] in ('감가상각비', '무형자산상각비'):
            return '상각비'
        return ACT2ITEM.get(r['활동분류(대분류)'], '기타관리경비')

    df['보고항목'] = df.apply(_보고항목, axis=1)
    if 연도힌트:
        df = df[df['년'] == 연도힌트]
    return df.reset_index(drop=True)


def 월간보고입력읽기(xls):
    """실적파일의 「월간보고_입력」 시트를 읽습니다. 없으면 빈 값으로 돌려줍니다."""
    빈값 = {'단일': {}, '월별': {}, '진척': {}, '서술': {}, '고객목표': {}}
    if '월간보고_입력' not in xls.sheet_names:
        return 빈값
    try:
        d = pd.read_excel(xls, sheet_name='월간보고_입력', header=None)
    except Exception:
        return 빈값
    구역 = None
    for _, row in d.iterrows():
        vals = list(row.values)
        라벨 = str(vals[1]).strip() if len(vals) > 1 and pd.notna(vals[1]) else ''
        if not 라벨 or 라벨.lower() == 'nan':
            continue
        if 라벨.startswith('■'):
            t = 라벨
            구역 = ('단일' if '기본' in t else '월별' if '월별' in t else
                    '고객목표' if '고객' in t else '진척' if '진척' in t else
                    '서술' if '서술' in t else None)
            continue
        if 구역 is None or 라벨.startswith('※') or 라벨 in ('항목', '구분', '고객사'):
            continue
        v2 = vals[2] if len(vals) > 2 else None
        if 구역 == '월별':
            빈값['월별'][라벨] = [float(vals[2 + i]) if len(vals) > 2 + i
                                  and isinstance(vals[2 + i], (int, float))
                                  and pd.notna(vals[2 + i]) else 0.0 for i in range(12)]
        elif 구역 == '서술':
            빈값['서술'][라벨] = '' if pd.isna(v2) else str(v2)
        else:
            if pd.isna(v2):
                continue
            try:
                빈값[구역][라벨] = float(v2)
            except (TypeError, ValueError):
                빈값[구역][라벨] = v2
    return 빈값


def _매핑표만들기(m):
    """BS_IS_매핑 시트 → 계정 영문명으로 찾아 쓰는 사전. 원장읽기와 엑셀작성이 함께 씁니다."""
    m = m.copy()
    m.columns = [str(c).strip() for c in m.columns]
    영문col = [c for c in m.columns if '영문' in c or 'account' in c.lower()][0]
    bsiscol = [c for c in m.columns if 'BS' in c][0]
    korcol = [c for c in m.columns if '한글' in c][0]
    key = m[영문col].astype(str).str.strip().str.lower()
    return {'bsis': dict(zip(key, m[bsiscol])), 'kor': dict(zip(key, m[korcol])),
            'table': m.rename(columns={영문col: '계정(영문)', bsiscol: 'BS/IS',
                                       korcol: '계정과목(한글)'})}


def _기초현금계산(raw, 매핑):
    """원장 원본(머리글 5행)에서 현금·예금 계정의 Beginning Balance 합계를 구합니다.

    Beginning Balance 줄은 거래가 아니라 기초잔액이라 _표준화()에서 걸러지므로
    여기서 따로 읽습니다. 계정 이름은 A열에 한 번만 적혀 있어 ffill 로 채웁니다.
    """
    현금계정 = {k for k, v in 매핑['kor'].items()
                if 매핑['bsis'].get(k) == 'BS' and isinstance(v, str)
                and ('현금' in v or '예금' in v)}
    try:
        if 'Distribution account' not in [str(c) for c in raw.columns]:
            return 0.0
        그룹 = raw.iloc[:, 0].astype('object').ffill().astype(str).str.strip().str.lower()
        계정 = raw['Distribution account'].astype(str).str.strip()
        잔액 = pd.to_numeric(raw.get('Balance'), errors='coerce').fillna(0)
        return float(잔액[계정.eq('Beginning Balance') & 그룹.isin(현금계정)].sum())
    except Exception:
        return 0.0



def _IS밑자료(xls, 매핑):
    """연결패키지 손익계산서에 넣을 「국문 계정과목별 금액」과, 계정(영문)별 금액.

    당기 : 26년 원장_raw 1~보고월 (수익은 대변−차변, 그 밖은 차변−대변)
    전기 : 25년 원장_raw 한 해 전체
    """
    def 한해(시트):
        if 시트 not in xls.sheet_names:
            return {}, {}
        raw = pd.read_excel(xls, sheet_name=시트, header=4)
        이름 = _글자칸(raw, 'Distribution account')
        차 = pd.to_numeric(raw.get('Debit'), errors='coerce').fillna(0)
        대 = pd.to_numeric(raw.get('Credit'), errors='coerce').fillna(0)
        날 = pd.to_datetime(raw['Transaction date'], errors='coerce')
        과목별, 계정별 = {}, {}
        for i in range(len(raw)):
            n = 이름[i]
            if not n or n == 'Beginning Balance' or pd.isna(날.iat[i]):
                continue
            키 = n.strip().lower()
            if 매핑['bsis'].get(키) != 'IS':
                continue
            과목 = str(매핑['kor'].get(키, '')).strip()
            매출 = 과목 in ('제품매출', '기타매출', '상품매출', '용역매출', '잡이익', '외환차익', '외화환산이익')
            v = (float(대.iat[i]) - float(차.iat[i])) if 매출 else (float(차.iat[i]) - float(대.iat[i]))
            계정별[n] = 계정별.get(n, 0.0) + v
            if 과목:
                과목별[과목] = 과목별.get(과목, 0.0) + v
        return 과목별, 계정별

    당기과목, 당기계정 = 한해('26년 원장_raw')
    전기과목, 전기계정 = 한해('25년 원장_raw')
    return {'당기': 당기과목, '전기': 전기과목,
            '당기계정': 당기계정, '전기계정': 전기계정}


def _영문BS자료(xls, 매핑):
    """퀵북 계정(영문) 이름 그대로의 재무상태표 — 전기 · 당기 · 증감."""
    전기, 차례 = {}, []
    if '26년 BS기초' in xls.sheet_names:
        d = pd.read_excel(xls, sheet_name='26년 BS기초', header=None)
        for _i, 줄 in d.iterrows():
            n = 줄.iloc[0]
            if n is None or (isinstance(n, float) and pd.isna(n)):
                continue
            n = str(n).strip()
            if not n:
                continue
            v = pd.to_numeric(pd.Series([줄.iloc[1]]), errors='coerce').fillna(0).iat[0]
            if n not in 전기:
                차례.append(n)
            전기[n] = 전기.get(n, 0.0) + float(v)

    raw = pd.read_excel(xls, sheet_name='26년 원장_raw', header=4)
    머리 = _글자칸(raw, raw.columns[0])
    그룹, 지금 = [], ''
    for v in 머리:
        if v:
            지금 = v
        그룹.append(지금)
    이름 = _글자칸(raw, 'Distribution account')
    차 = pd.to_numeric(raw.get('Debit'), errors='coerce').fillna(0)
    대 = pd.to_numeric(raw.get('Credit'), errors='coerce').fillna(0)
    잔 = pd.to_numeric(raw.get('Balance'), errors='coerce').fillna(0)
    날 = pd.to_datetime(raw['Transaction date'], errors='coerce')
    기초, 증감 = {}, {}
    for i in range(len(raw)):
        if 이름[i] == 'Beginning Balance':
            기초[그룹[i]] = 기초.get(그룹[i], 0.0) + float(잔.iat[i])
        elif 이름[i] and pd.notna(날.iat[i]):
            증감[이름[i]] = 증감.get(이름[i], 0.0) + float(차.iat[i]) - float(대.iat[i])
    당기 = {}
    for 계정 in set(기초) | set(증감):
        키 = 계정.strip().lower()
        과목 = str(매핑['kor'].get(키, '')).strip()
        if 매핑['bsis'].get(키) != 'BS':
            continue
        b, m = 기초.get(계정, 0.0), 증감.get(계정, 0.0)
        if 과목 in BS자산칸부채:
            당기[계정] = m - b
        elif 과목 in BS부채자본:
            당기[계정] = b - m
        else:
            당기[계정] = b + m
    for 계정 in 당기:
        if 계정 not in 전기:
            차례.append(계정)
    # 퀵북 표 차례대로 그릴 때 이름으로 찾아 쓰는 전기 값 (빈 칸은 빈 칸 그대로)
    전기맞춤 = {}
    if '26년 BS기초' in xls.sheet_names:
        d = pd.read_excel(xls, sheet_name='26년 BS기초', header=None)
        for _i, 줄 in d.iterrows():
            n = 줄.iloc[0]
            if n is None or (isinstance(n, float) and pd.isna(n)):
                continue
            키 = _맞춤이름(n)
            if not 키:
                continue
            v = pd.to_numeric(pd.Series([줄.iloc[1]]), errors='coerce').iat[0]
            if pd.notna(v):
                전기맞춤[키] = float(v)
            else:
                전기맞춤.setdefault(키, None)
    return {'차례': 차례, '전기': 전기, '당기': 당기, '전기맞춤': 전기맞춤}


꼬리줄 = re.compile(r'(accrual|cash)\s+basis|\w+day,\s', re.I)


def _총계이름(n):
    """「Total for Income」 · 「Total Income」 → 'income' (아니면 None)."""
    m = re.match(r'^\s*total\s+(?:for\s+)?(.+?)\s*$', str(n or ''), re.I)
    return re.sub(r'\s+', ' ', m.group(1)).lower() if m else None


def _맞춤이름(n):
    return re.sub(r'^total for\s+', 'total ',
                  re.sub(r'\s+', ' ', str(n or '')).strip().lower())


@st.cache_data(show_spinner=False)
def 퀵북표읽기(_도장, 종류):
    """보관해 둔 원장의 「○월(BS)」·「○월(IS)」 시트를 차례·들여쓰기 그대로 읽습니다.

      · 줄 차례는 퀵북이 뽑아 준 그대로입니다 (구역 머리글·Total 줄 포함)
      · 들여쓰기는 「뒤에 Total 줄이 있는 이름 = 구역을 여는 줄」 로 되살립니다
    """
    d = pd.read_excel(원장보관, sheet_name=None, header=None)
    시트, 달 = None, 0
    for s in d:
        m = re.search(r'(\d+)\s*월\s*\(\s*(BS|IS)\s*\)', str(s), re.I)
        if m and m.group(2).upper() == 종류:
            시트, 달 = s, int(m.group(1))
    if 시트 is None:
        return None
    표 = d[시트]
    제목, 이름들, 금액들 = [], [], []
    for i, 줄 in 표.iterrows():
        n = 줄.iloc[0]
        n = '' if n is None or (isinstance(n, float) and pd.isna(n)) else str(n).strip()
        v = pd.to_numeric(pd.Series([줄.iloc[1] if len(줄) > 1 else None]),
                          errors='coerce').iat[0]
        if i < 5:
            if n and n.lower() != 'total':
                제목.append(n)
            continue
        if not n or 꼬리줄.search(n):
            continue
        이름들.append(n)
        금액들.append(None if pd.isna(v) else float(v))
    남은 = {}
    for n in 이름들:
        t = _총계이름(n)
        if t:
            남은[t] = 남은.get(t, 0) + 1
    깊이, 쌓기 = [], []
    for n in 이름들:
        t = _총계이름(n)
        키 = re.sub(r'\s+', ' ', n).lower()
        if t and t in 쌓기:
            while 쌓기 and 쌓기[-1] != t:
                쌓기.pop()
            쌓기.pop()
            깊이.append(len(쌓기))
        else:
            깊이.append(len(쌓기))
            if 남은.get(키, 0) > 0 and 키 not in 쌓기:
                쌓기.append(키)
                남은[키] -= 1
    while len(제목) < 3:
        제목.append('')
    return {'제목': 제목, '시트': 시트, '월': 달,
            '줄': [{'이름': n, '깊이': g, '당기': v, '총계': bool(_총계이름(n))}
                   for n, g, v in zip(이름들, 깊이, 금액들)]}


# 「26년 자금」 시트의 활동성항목 → 현금흐름 표의 줄
예상항목줄 = {'매출채권 회수': 'AR', '기타입금': 'ETCIN', '매입채무 지급': 'AP',
              '급여': 'PAY', '인건비성 항목': 'PAY', '기타 비용': 'OTHOP',
              '설비 투자': 'CAPEX', '유상증자': 'EQ', '차입증감': '차입증감',
              '리스료 지급': '리스', '환율변동 효과': 'FX'}
# 돈이 「들어오는」 항목 — 나머지는 나가는 것으로 봅니다.
#  ※ 차변/대변 어느 칸에 적으셔도 항목 이름으로 방향을 정합니다.
예상입금항목 = {'매출채권 회수', '기타입금', '유상증자', '차입증감'}


@st.cache_data(show_spinner=False)
def 예상자금읽기(데이터바이트):
    """「26년 자금」 시트의 2번(예상 출금) · 3번(예상 입금)을 월별로 모읍니다.

      · 칸 자리를 미리 정해 두지 않고 머리글(일자·차변·대변·활동성항목)을 보고 찾습니다.
        그래서 「기말」 칸을 빼시거나 자리를 옮기셔도 그대로 읽습니다.
      돌려주는 값 : {'2': {줄이름: {월: 금액}}, '3': {…}} — 들어온 돈 +, 나간 돈 −
        2번은 클로드가 채우는 칸, 3번은 담당자가 손으로 채우는 칸이라 따로 담습니다.
    """
    try:
        xls = pd.ExcelFile(io.BytesIO(데이터바이트))
        시트 = next((s for s in xls.sheet_names if '자금' in str(s)), None)
        if 시트 is None:
            return None
        d = pd.read_excel(xls, sheet_name=시트, header=None)
    except Exception:
        return None

    def 글자(i, j):
        v = d.iat[i, j] if 0 <= i < len(d) and 0 <= j < d.shape[1] else None
        return '' if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()

    머리행 = next((i for i in range(min(12, len(d)))
                   if any(글자(i, j) == '일자' for j in range(d.shape[1]))), None)
    if 머리행 is None:
        return {'2': {}, '3': {}}
    일자열 = [j for j in range(d.shape[1]) if 글자(머리행, j) == '일자']
    모음 = {'2': {}, '3': {}}
    for k, 첫 in enumerate(일자열):
        끝 = 일자열[k + 1] if k + 1 < len(일자열) else d.shape[1]
        자리 = {}
        for j in range(첫, 끝):
            이름 = 글자(머리행, j)
            if 이름 in ('차변', '대변', '활동성항목') and 이름 not in 자리:
                자리[이름] = j
        # 이 칸이 몇 번 표인지 (머리글 위쪽에서 「2.」·「3.」 을 찾습니다)
        번호 = ''
        for i in range(머리행):
            for j in range(max(0, 첫 - 1), 끝):
                m = re.match(r'^\s*([123])\s*\.', 글자(i, j))
                if m:
                    번호 = m.group(1)
        if 번호 not in ('2', '3') or '활동성항목' not in 자리:
            continue                      # 1번(실제)은 원장에서 이미 읽었습니다
        날 = pd.to_datetime(d.iloc[:, 첫], errors='coerce')
        차 = (pd.to_numeric(d.iloc[:, 자리['차변']], errors='coerce').fillna(0)
              if '차변' in 자리 else pd.Series(0.0, index=d.index))
        대 = (pd.to_numeric(d.iloc[:, 자리['대변']], errors='coerce').fillna(0)
              if '대변' in 자리 else pd.Series(0.0, index=d.index))
        이름칸 = d.iloc[:, 자리['활동성항목']]
        for i in range(len(d)):
            if pd.isna(날.iat[i]):
                continue
            항목 = 이름칸.iat[i]
            항목 = '' if 항목 is None or (isinstance(항목, float) and pd.isna(항목)) \
                else str(항목).strip()
            줄 = 예상항목줄.get(항목, 'OTHOP')
            # 차변·대변 어느 칸에 넣으셔도 됩니다 — 방향은 활동성항목으로 정합니다
            금액 = float(차.iat[i]) + float(대.iat[i])
            v = 금액 if 항목 in 예상입금항목 else -금액
            if not v:
                continue
            칸모음 = 모음[번호]
            칸모음.setdefault(줄, {})
            m = int(날.iat[i].month)
            칸모음[줄][m] = 칸모음[줄].get(m, 0.0) + v
    return 모음


def 예상현금모형(매출L, 원가L, 판관L, 상각L, 자금표, 보고월, 끝월=12):
    """앞으로 올 달(보고월+1 ~ 12월)의 현금 증감을 실적에서 그대로 뽑아 추정합니다.

    ■ 왜 필요한가
      예전에는 「월별 실적집계」 관리비 줄의 월평균만 빼는 방식이었습니다. 그런데
      관리비 안에 ① SR&ED 정부보조금처럼 한 번 들어오고 마는 돈과 ② 상각비처럼
      현금이 아예 나가지 않는 비용이 섞여 있어, 앞으로 매달 돈이 들어오는 것처럼
      잡히고 매출·매출원가는 아예 빠져 있었습니다. 그래서 계산을 다시 세웁니다.

    ■ 쓰는 가정은 딱 셋 — 셋 다 1~보고월 실적을 그대로 나눈 값입니다
      · 매출액   = 1~보고월 실제 매출 ÷ 보고월           (월평균)
      · 매출원가 = 예상 매출액 × 1~보고월 실제 원가율     (매출원가 ÷ 매출액)
      · 판관비   = (1~보고월 실제 판관비 − 상각비) ÷ 보고월
                   → 상각비는 현금이 나가지 않아 뺍니다. SR&ED 보조금은 판관비
                     차감분이라 애초에 여기 안 들어옵니다(일회성이라 뺍니다).
      운전자본(매출채권·매입채무·재고)은 지금 수준 그대로라고 봅니다. 그래서
      「매출채권 회수 = 그 달 매출」, 「매입채무 지급 = 그 달 매출원가」 입니다.
      설비 투자 · 유상증자 · 차입은 회사가 정하는 일이라 넣지 않았습니다
      (담당자가 「26년 자금」 3번 표에 넣으면 그대로 덧붙습니다).

    돌려주는 값 : ({줄이름: {월: 금액}}, 가정메모) — 들어온 돈 +, 나간 돈 −
    """
    달수 = max(int(보고월), 1)
    예상월 = list(range(보고월 + 1, 끝월 + 1))       # 13 = 다음 해 1월 … 18 = 다음 해 6월
    매출합 = float(sum(매출L))
    원가합 = float(sum(원가L))
    판관합 = float(sum(판관L))
    상각합 = float(sum(상각L))
    월매출 = 매출합 / 달수
    원가율 = (원가합 / 매출합) if abs(매출합) > 1e-9 else 0.0
    월원가 = 월매출 * 원가율
    월판관 = max((판관합 - 상각합) / 달수, 0.0)
    # 급여는 「현금이 실제로 나간 급여」 월평균을 씁니다 (발생액이 아니라 지급액)
    월급여 = 0.0
    if 자금표 is not None and 'PAY' in getattr(자금표, 'index', []):
        월급여 = abs(float(sum(float(자금표.loc['PAY', m])
                               for m in range(1, 보고월 + 1)))) / 달수
    월급여 = min(월급여, 월판관)
    월기타 = max(월판관 - 월급여, 0.0)

    예상 = {}

    def 깔기(줄, 값):
        if abs(값) < 0.005:
            return
        예상.setdefault(줄, {})
        for 월 in 예상월:
            예상[줄][월] = 예상[줄].get(월, 0.0) + 값

    깔기('AR', +월매출)          # 매출채권 회수 — 그 달 매출이 그 달 안에 들어온다고 봅니다
    깔기('AP', -월원가)          # 매입채무 지급 — 그 달 매출원가만큼 나간다고 봅니다
    깔기('PAY', -월급여)         # 급여
    깔기('OTHOP', -월기타)       # 나머지 현금 판관비

    가정 = {'월매출': 월매출, '원가율': 원가율, '월원가': 월원가,
            '월판관': 월판관, '월급여': 월급여, '월기타': 월기타,
            '월상각': 상각합 / 달수, '월판관총': 판관합 / 달수,
            '월순증감': 월매출 - 월원가 - 월판관, '달수': 달수, '예상월': 예상월}
    return 예상, 가정


@st.cache_data(show_spinner='원장을 읽는 중입니다...')
def 원장읽기(데이터바이트):
    xls = pd.ExcelFile(io.BytesIO(데이터바이트))
    시트 = xls.sheet_names
    매핑표 = _매핑표만들기(pd.read_excel(xls, sheet_name='BS_IS_매핑'))

    def _read(sheet, 연도):
        if sheet not in 시트:
            return None
        for header in (0, 4):
            try:
                d = pd.read_excel(xls, sheet_name=sheet, header=header)
            except Exception:
                continue
            cols = [str(c) for c in d.columns]
            if any(c in cols for c in ('계정영문', 'Distribution account')):
                out = _표준화(d, 매핑표, 연도)
                if out is not None and len(out):
                    return out
        return None

    g26 = _read('26년 원장_raw', None)
    g25 = _read('25년 원장_raw', None)
    if g26 is None:
        raise ValueError('26년 원장_raw 시트를 읽을 수 없습니다.')
    당해 = int(g26['년'].mode().iat[0])
    g26 = g26[g26['년'] == 당해].reset_index(drop=True)
    if g25 is not None:
        g25 = g25[g25['년'] == 당해 - 1].reset_index(drop=True)
    else:
        g25 = g26.iloc[0:0].copy()

    # ── 현금 계정의 기초잔액
    try:
        기초현금 = _기초현금계산(pd.read_excel(xls, sheet_name='26년 원장_raw', header=4), 매핑표)
    except Exception:
        기초현금 = 0.0

    입력값 = 월간보고입력읽기(xls)
    # ── 연결패키지 재무상태표에 쓸 밑자료 (전기 잔액 · 당기 계정별 잔액)
    try:
        BS자료 = _BS밑자료(xls, 매핑표)
        BS자료['IS'] = _IS밑자료(xls, 매핑표)
        BS자료['영문BS'] = _영문BS자료(xls, 매핑표)
    except Exception:
        BS자료 = None
    # ── 기준 자료가 「연결패키지」(BS·IS 시트가 든 파일)이면 그 시트의 숫자를 그대로 씁니다
    #    (패키지 BS·IS 는 원장 SUMIFS 수식 + 본사 기초 — 화면이 패키지와 1원도 어긋나지 않게)
    if BS자료 is not None and 'BS' in 시트 and 'IS' in 시트:
        try:
            패 = _패키지BS밑자료(xls)
            if 패['BS당기'] and 패['IS당기']:
                BS자료['전기'], BS자료['당기'] = 패['BS전기'], 패['BS당기']
                BS자료['IS']['전기'], BS자료['IS']['당기'] = 패['IS전기'], 패['IS당기']
                BS자료['패키지시트'] = True
        except Exception:
            pass
        try:   # 환율 칸 (BS H5 전기말 · BS M5 당기말 · IS M5 평균) — CTKonnect 패키지 값을 화면 기본값으로
            _bs = pd.read_excel(xls, sheet_name='BS', header=None, nrows=6)
            _num = lambda v: float(v) if isinstance(v, (int, float)) and not pd.isna(v) and v > 0 else None
            BS자료['환율'] = {'전기': _num(_bs.iat[4, 7]) if _bs.shape[1] > 7 else None,
                              '당기': _num(_bs.iat[4, 12]) if _bs.shape[1] > 12 else None}
        except Exception:
            pass
        try:
            BS자료['수불원가'] = 수불원가읽기(xls)
            BS자료['품목군'] = 수불품목군읽기(xls)
            BS자료['원가절감'] = 조직운영읽기(xls)          # CTKonnect : 「조직운영안」 시트 하나
            BS자료['판관과제'] = None
            BS자료['총액'] = 총액읽기(xls)
            BS자료['기타사항'] = 기타사항읽기(xls)
        except Exception:
            BS자료['수불원가'] = None
    return g26, g25, 당해, 매핑표, 기초현금, 입력값, BS자료


# ── 원가 기준 손익 — 재고수불부(v61)와 같은 정의 (2026-10-27) ──────────────
#   직접노무비(생산인력) = 공장 생산직 7명 Wages + 개인 계약직(Subcontractors, CSEC 제외)
#   제조간접비(감가상각비 등) = 제조 상각 3계정 + CSEC 전기 컨설팅 (노무비는 들어가지 않음)
#   직접재료비 · 외주가공비 = 재고수불부 값 (패키지 「요약보고」 ①-2 표에서 읽음)
생산인력 = {'jitendra panda', 'chungyeon cho', 'william kim', 'chi yiu weng',
            'shiu fung aric li', 'thien an nguyen', 'gia bao chu'}
제조상각계정 = {'depreciation-mfg', 'depreciation - lease improvement',
                'depreciation - tooling and others'}
원가기준라벨 = '원가 기준 (재고수불부)'


def 원가기준표(IS, 매출L, 원가L, 판관L, SREDL, months, 수불원가):
    """원가 기준(재고수불부) 손익 줄들을 만듭니다. 비용은 양수.

    · 직접재료비 · 외주가공비 : 수불부 값(요약보고 ①-2). 수불부가 없는 달은
      「가정」 — 수불부가 있는 달의 평균 원가율(재료비 ÷ 매출, 외주 ÷ 매출) × 그 달 매출.
    · 직접노무비 = 생산직 급여 + 개인 계약직(CSEC 제외), 제조간접비 = 제조 상각 3계정 + CSEC (원장)
    · 판관비 = 원장 판관비 − 직접노무비 − 제조간접비 (상위 5개 계정 + 기타 + SR&ED 차감)
    """
    def 월별(mask, 부호=-1):
        g = IS[mask].groupby('월')['금액'].sum().reindex(months, fill_value=0.0) * 부호
        return [float(g.get(m, 0.0)) for m in months]
    영문 = IS['계정영문'].astype(str).str.strip().str.lower()
    이름열 = (IS['Name'].astype(str).str.strip().str.lower() if 'Name' in IS.columns
              else pd.Series('', index=IS.index))
    계약직 = 영문.eq('subcontractors') & ~이름열.str.contains('csec')
    시섹 = 영문.eq('subcontractors') & 이름열.str.contains('csec')
    생산급여 = 영문.eq('wages') & 이름열.isin(생산인력)
    노무L = 월별(생산급여 | 계약직)
    간접L = 월별(영문.isin(제조상각계정) | 시섹)
    수불재료 = 수불원가.get('직접재료비', {}) if 수불원가 else {}
    수불외주 = 수불원가.get('외주가공비', {}) if 수불원가 else {}
    있는달 = [i for i, m in enumerate(months) if m in 수불재료]
    기준매출 = sum(매출L[i] for i in 있는달)
    재료율 = (sum(float(수불재료[months[i]]) for i in 있는달) / 기준매출) if 기준매출 else 0.0
    외주율 = (sum(float(수불외주.get(months[i], 0.0)) for i in 있는달) / 기준매출) if 기준매출 else 0.0
    재료L, 외주L, 가정월 = [], [], []
    for i, m in enumerate(months):
        if m in 수불재료:
            재료L.append(float(수불재료[m])); 외주L.append(float(수불외주.get(m, 0.0)))
        else:
            재료L.append(매출L[i] * 재료율); 외주L.append(매출L[i] * 외주율); 가정월.append(m)
    n = len(months)
    변동L = [재료L[i] + 외주L[i] + 노무L[i] for i in range(n)]
    매출이익L = [매출L[i] - 변동L[i] - 간접L[i] for i in range(n)]
    판관원가L = [판관L[i] - 노무L[i] - 간접L[i] for i in range(n)]
    제외 = 생산급여 | 계약직 | 시섹 | 영문.isin(제조상각계정)
    판관조정 = IS['분류'].eq('판관비') & ~제외
    계정별 = IS[판관조정].groupby('계정과목')['금액'].sum().mul(-1).sort_values(ascending=False)
    상위5 = [str(a) for a in 계정별.index[:5] if str(a) != 'nan']
    세부 = [(이름, 월별(판관조정 & IS['계정과목'].eq(이름))) for 이름 in 상위5]
    세부.append(('기타', [판관원가L[i] - sum(v[1][i] for v in 세부) for i in range(n)]))
    SRED줄 = [-v for v in SREDL]
    세부.append(('SR&ED 차감', SRED줄))
    판관순L = [판관원가L[i] + SRED줄[i] for i in range(n)]
    영익L = [매출이익L[i] - 판관순L[i] for i in range(n)]
    return {'재료': 재료L, '외주': 외주L, '노무': 노무L, '간접': 간접L, '변동': 변동L,
            '매출이익': 매출이익L, '판관': 판관원가L, '세부': 세부, '판관순': 판관순L,
            '영익': 영익L, '가정월': 가정월, '재료율': 재료율, '외주율': 외주율,
            '수불월': [months[i] for i in 있는달], '상위5': 상위5}


def 수불원가읽기(xls):
    """기준 자료의 「요약보고」 ①-2 표(재고수불부 기준 월별)에서 직접재료비 · 외주가공비 등을 읽습니다.
    돌려주는 값 : {'직접재료비': {월: 금액}, '외주가공비': {…}, …} — 값이 없는 달은 빠집니다."""
    if '요약보고' not in xls.sheet_names:
        return None
    d = pd.read_excel(xls, sheet_name='요약보고', header=None)
    시작 = None
    for i in range(len(d)):
        v = d.iat[i, 1] if d.shape[1] > 1 else None
        if isinstance(v, str) and v.strip().startswith('①-2'):
            시작 = i
            break
    if 시작 is None:
        return None
    머리 = 시작 + 1
    월칸 = {}
    for c in range(2, d.shape[1]):
        v = pd.to_numeric(pd.Series([d.iat[머리, c]]), errors='coerce').iat[0]
        if pd.notna(v) and 1 <= int(v) <= 12 and float(v) == int(v):
            월칸[int(v)] = c
    out = {}
    for i in range(머리 + 1, min(머리 + 40, len(d))):
        라벨 = d.iat[i, 1]
        if 라벨 is None or (isinstance(라벨, float) and pd.isna(라벨)):
            break
        라벨 = str(라벨).strip()
        줄 = {}
        for m, c in 월칸.items():
            v = pd.to_numeric(pd.Series([d.iat[i, c]]), errors='coerce').iat[0]
            if pd.notna(v):
                줄[m] = float(v)
        out[라벨] = 줄
    return out


품목군목록 = ['Resin', 'Straw', 'Cutlery']


def 수불품목군읽기(xls):
    """기준 자료의 「요약보고(수불부 8월)」 ② 제품군별 · ④ 거래처별 표 → 품목군 배분 기준.
    돌려주는 값 : {'매출': {군: 금액}, '재료': {...}, '외주': {...}, '레진거래처': set} — 없으면 None."""
    시트 = next((n for n in xls.sheet_names if str(n).startswith('요약보고(수불부')), None)
    if 시트 is None:
        return None
    d = pd.read_excel(xls, sheet_name=시트, header=None)

    def 글(i, c):
        v = d.iat[i, c] if c < d.shape[1] else None
        return '' if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()

    시작 = next((i for i in range(len(d)) if 글(i, 1).startswith('② 제품군')), None)
    if 시작 is None:
        return None
    머리 = 시작 + 1
    칸 = {글(머리, c): c for c in range(2, min(d.shape[1], 7)) if 글(머리, c) in 품목군목록}
    if len(칸) < 3:
        return None
    out = {'매출': {}, '재료': {}, '외주': {}, '레진거래처': set()}
    for i in range(머리 + 1, min(머리 + 20, len(d))):
        라벨 = 글(i, 1).replace(' ', '')
        키 = ('매출' if 라벨.startswith('매출액') else '재료' if 라벨.startswith('직접재료비')
              else '외주' if 라벨.startswith('외주가공비') else None)
        if 키:
            out[키] = {g: float(pd.to_numeric(pd.Series([d.iat[i, c]]), errors='coerce').fillna(0).iat[0])
                       for g, c in 칸.items()}
    # ④ 거래처별 — 비고에 「레진판매」 인 거래처
    시작4 = next((i for i in range(len(d)) if 글(i, 1).startswith('④')), None)
    if 시작4 is not None:
        머리4 = 시작4 + 1
        비고칸 = next((c for c in range(2, d.shape[1]) if 글(머리4, c) == '비고'), None)
        if 비고칸 is not None:
            for i in range(머리4 + 1, min(머리4 + 80, len(d))):
                if not 글(i, 1) or 글(i, 1) == '합계':
                    break
                if '레진' in 글(i, 비고칸):
                    out['레진거래처'].add(글(i, 1).lower())
    return out if out['매출'] else None


def 품목군손익표(IS, months, 매출L, C, 품목군):
    """원가 기준 손익을 품목군(Resin · Straw · Cutlery)으로 나눕니다 (약식).
    · 매출액 : 원장 거래처별 월 매출 — 레진 거래처(수불부 ④ 「레진판매」) → Resin,
               그 밖의 거래처 → Straw · Cutlery 를 수불부 1~8월 매출 비율로 나눔
    · 직접재료비 · 외주가공비 : 품목군 매출 × 수불부 원가율(재료/매출 · 외주/매출) → 그 달 합계에 맞춰 비례 조정
    · 직접노무비 · 고정비(제조간접비) · 판관비 : 그 달 품목군 매출 비중으로 배분 (수불부와 같은 방식)
    돌려주는 값 : {군: {'매출': [...], '변동비': [...], '고정비': [...], '판관비': [...], '영업이익': [...]}}"""
    n = len(months)
    매출행 = IS[IS['분류'].eq('매출')]
    레진 = 매출행['Name'].fillna('').astype(str).str.strip().str.lower().isin(품목군['레진거래처'])
    월합 = lambda 조건: [float(매출행[조건 & 매출행['월'].eq(m)]['금액'].sum()) for m in months]
    레진L, 그외L = 월합(레진), 월합(~레진)
    s, c = 품목군['매출'].get('Straw', 0.0), 품목군['매출'].get('Cutlery', 0.0)
    빨대비 = s / (s + c) if (s + c) else 1.0
    매출 = {'Resin': 레진L, 'Straw': [v * 빨대비 for v in 그외L], 'Cutlery': [v * (1 - 빨대비) for v in 그외L]}
    # IS 매출 합계와 맞춥니다 (거래처 없는 매출 등은 비율대로)
    for i in range(n):
        합, 전체 = sum(매출[g][i] for g in 품목군목록), 매출L[i]
        if 합:
            for g in 품목군목록:
                매출[g][i] *= 전체 / 합

    def 율(키, g):
        m = 품목군['매출'].get(g, 0.0)
        return (품목군[키].get(g, 0.0) / m) if m else 0.0

    out = {g: {'매출': 매출[g]} for g in 품목군목록}
    for i in range(n):
        총매출 = sum(매출[g][i] for g in 품목군목록)
        for 키, 합계 in (('재료', C['재료'][i]), ('외주', C['외주'][i])):
            기준 = {g: 매출[g][i] * 율(키, g) for g in 품목군목록}
            합 = sum(기준.values())
            for g in 품목군목록:
                out[g].setdefault(키, []).append(합계 * 기준[g] / 합 if 합 else 0.0)
        for g in 품목군목록:
            비 = 매출[g][i] / 총매출 if 총매출 else 0.0
            out[g].setdefault('노무', []).append(C['노무'][i] * 비)
            out[g].setdefault('고정비', []).append(C['간접'][i] * 비)
            out[g].setdefault('판관비', []).append(C['판관'][i] * 비)
    for g in 품목군목록:
        o = out[g]
        o['변동비'] = [o['재료'][i] + o['외주'][i] + o['노무'][i] for i in range(n)]
        o['영업이익'] = [o['매출'][i] - o['변동비'][i] - o['고정비'][i] - o['판관비'][i] for i in range(n)]
    return out


def _효율화블록읽기(d):
    """「효율화」 시트 — 월별 블록 모양 (제목 「2026년 9월 — 내부 효율화」 + 머리 줄 + 입력 줄)."""
    out, 지금, 열 = [], None, {}

    def 값(i, c):
        if c is None or c >= d.shape[1]:
            return None
        v = d.iat[i, c]
        return None if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == '' else v

    for i in range(len(d)):
        줄 = [str(v).strip() if v is not None and not (isinstance(v, float) and pd.isna(v)) else ''
              for v in d.iloc[i].tolist()]
        첫 = next((t for t in 줄 if t), '')
        m = re.match(r'^(20\d\d)\s*년\s*(\d{1,2})\s*월', 첫)
        if m:
            지금 = (int(m.group(1)), int(m.group(2)))
            continue
        if '효율화 항목' in 줄 or '개선과제' in 줄 or '과제 구분' in 줄:
            열 = {t: c for c, t in enumerate(줄) if t}
            continue
        if 지금 is None or not 열:
            continue
        찾 = lambda 앞: next((c for t, c in 열.items() if t.startswith(앞)), None)
        새양식 = 찾('과제 구분') is not None
        if 새양식:            # 2026-10-27 양식 : 과제 구분 · 업체명 · 신규/기존 · 실행 내용 · 목표 금액 · 목표 시점
            구분, 업체, 신규 = 값(i, 찾('과제 구분')), 값(i, 찾('업체명')), 값(i, 찾('신규'))
            항목, 금액값, 시점 = 값(i, 찾('실행 내용')), 값(i, 찾('절감 금액') if 찾('절감 금액') is not None else 찾('목표 금액')), 값(i, 찾('목표 시점'))
            if all(v is None for v in (구분, 업체, 항목, 금액값)):
                continue
            if hasattr(시점, 'year'):
                시점 = f'{시점.year}-{시점.month:02d}'
            누락 = [n for n, v in (('구분', 구분), ('업체명', 업체), ('신규/기존', 신규), ('실행 내용', 항목),
                                    ('절감 금액', 금액값), ('목표 시점', 시점)) if v is None]
            금액 = pd.to_numeric(pd.Series([금액값]), errors='coerce').iat[0]
            out.append({'연': 지금[0], '월': 지금[1], '구분': str(구분 or ''), '업체': str(업체 or ''),
                        '신규': str(신규 or ''), '항목': str(항목 or ''), '내용': '', '시점': str(시점 or ''),
                        '금액': None if pd.isna(금액) else float(금액), '누락': 누락,
                        '상태': str(값(i, 찾('진행 상태')) or ''), '담당': str(값(i, 찾('담당')) or '')})
            continue
        항목, 내용 = 값(i, 찾('개선과제') if 찾('개선과제') is not None else 찾('효율화 항목')), 값(i, 찾('진행 내용'))
        if 항목 is None and 내용 is None:
            continue
        금액 = pd.to_numeric(pd.Series([값(i, 찾('월 개선'))]), errors='coerce').iat[0]
        out.append({'연': 지금[0], '월': 지금[1], '구분': str(값(i, 찾('구분')) or ''),
                    '항목': str(항목 or ''), '내용': str(내용 or ''),
                    '금액': None if pd.isna(금액) else float(금액),
                    '상태': str(값(i, 찾('진행 상태')) or ''), '담당': str(값(i, 찾('담당')) or '')})
    return out


def 기타사항읽기(xls, 시트='기타주요사항'):
    """「기타주요사항」 시트 — 달 구분 없는 한 목록. 마지막 머리줄(구분 · 주요 내용 …) 아래 줄들을 읽습니다
    (위쪽 「예시」 줄은 건너뜀). 시트가 없으면 None."""
    if 시트 not in xls.sheet_names:
        return None
    d = pd.read_excel(xls, sheet_name=시트, header=None, dtype=object)
    def 글(v):
        return '' if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    머리들 = [i for i in range(len(d)) if any(글(v).startswith('주요 내용') for v in d.iloc[i].tolist())]
    if not 머리들:
        return []
    머리 = 머리들[-1]
    열 = {글(v): c for c, v in enumerate(d.iloc[머리].tolist()) if 글(v)}
    찾 = lambda 앞: next((c for t, c in 열.items() if t.startswith(앞)), None)
    out = []
    for i in range(머리 + 1, len(d)):
        def 값(앞):
            c = 찾(앞)
            if c is None:
                return ''
            v = d.iat[i, c]
            if hasattr(v, 'year') and not isinstance(v, str):
                return f'{v.year}-{v.month:02d}'
            return 글(v)
        r = {'구분': 값('구분'), '내용': 값('주요 내용'), '상태': 값('진행 상태'), '기한': 값('기한'),
             '담당': 값('담당'), '비고': 값('비고')}
        if not any(r.values()):
            continue
        r['누락'] = [n for n, k in (('구분', '구분'), ('주요 내용', '내용'), ('진행 상태', '상태')) if not r[k]]
        out.append(r)
    return out


def 총액읽기(xls, 시트='요약보고'):
    """요약보고 「총액기준 월별 영업손익」 표 → {'중개매입': {월: 금액}, 'IOR대납': {…}, '거래액': {…}} (없으면 None)."""
    if 시트 not in xls.sheet_names:
        return None
    d = pd.read_excel(xls, sheet_name=시트, header=None)
    글 = lambda v: '' if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    시작 = next((i for i in range(len(d)) if '총액' in 글(d.iat[i, 1]) and '월별' in 글(d.iat[i, 1])), None)
    if 시작 is None:
        return None
    머리 = next((i for i in range(시작 + 1, min(시작 + 4, len(d))) if 글(d.iat[i, 1]) == '구분'), None)
    if 머리 is None:
        return None
    월칸 = {}
    for c in range(2, d.shape[1]):
        v = pd.to_numeric(pd.Series([d.iat[머리, c]]), errors='coerce').iat[0]
        if not pd.isna(v) and 1 <= int(v) <= 12 and float(v) == int(v):
            월칸[int(v)] = c
    out = {}
    for i in range(머리 + 1, min(머리 + 12, len(d))):
        이름 = 글(d.iat[i, 1])
        k = ('중개매입' if '중개 매입' in 이름 else 'IOR대납' if 'IOR' in 이름 and '대납' in 이름
             else '거래액' if '거래액' in 이름 else None)
        if k:
            out[k] = {m: float(pd.to_numeric(pd.Series([d.iat[i, c]]), errors='coerce').fillna(0).iat[0])
                      for m, c in 월칸.items()}
    return out if ('중개매입' in out and 'IOR대납' in out) else None


def 조직운영읽기(xls, 시트='조직운영안'):
    """「조직운영안」 시트 (2026-10-28 양식) — 월별 칸 : 제목 「2026년 8월 — 조직운영안」 + 머리줄 + 입력줄.
    머리 : 과제 구분 · 대상 업무 · 현재 (As-Is) · 변경안 (To-Be) · 기대 효과 (USD/월) · 목표 시점 · 진행 상태 · 담당자.
    캐나다 과제 칸과 같은 모양({구분, 항목, 내용, 금액 …})으로 돌려줍니다 — 항목 = 대상 업무, 내용 = 현재 → 변경안."""
    if 시트 not in xls.sheet_names:
        return None
    d = pd.read_excel(xls, sheet_name=시트, header=None, dtype=object)
    빈 = lambda v: v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ''
    out, 지금, 열 = [], None, {}
    for i in range(len(d)):
        줄 = ['' if 빈(v) else str(v).strip() for v in d.iloc[i].tolist()]
        첫 = next((t for t in 줄 if t), '')
        m = re.match(r'^(20\d\d)\s*년\s*(\d{1,2})\s*월', 첫)
        if m:
            지금 = (int(m.group(1)), int(m.group(2)))
            continue
        if '과제 구분' in 줄:
            열 = {t: c for c, t in enumerate(줄) if t}
            continue
        if 지금 is None or not 열:
            continue

        def 값(앞):
            c = next((c for t, c in 열.items() if t.startswith(앞)), None)
            v = None if c is None else d.iat[i, c]
            if 빈(v):
                return None
            if hasattr(v, 'year') and not isinstance(v, str):
                return f'{v.year}-{v.month:02d}'
            return v
        구분, 업무, 현재, 변경 = 값('과제 구분'), 값('대상 업무'), 값('현재'), 값('변경안')
        효과, 시점 = 값('기대 효과'), 값('목표 시점')
        if all(v is None for v in (구분, 업무, 현재, 변경, 효과)):
            continue
        누락 = [n for n, v in (('과제 구분', 구분), ('대상 업무', 업무), ('현재', 현재),
                                ('변경안', 변경), ('목표 시점', 시점)) if v is None]
        금액 = pd.to_numeric(pd.Series([효과]), errors='coerce').iat[0]
        out.append({'연': 지금[0], '월': 지금[1], '구분': str(구분 or ''), '업체': '', '신규': '',
                    '항목': str(업무 or ''),
                    '내용': (f'{현재 or "?"} → {변경 or "?"}' if (현재 or 변경) else ''),
                    '시점': str(시점 or ''), '금액': None if pd.isna(금액) else float(금액), '누락': 누락,
                    '상태': str(값('진행 상태') or ''), '담당': str(값('담당') or '')})
    return out


def 원가절감읽기(xls, 시트들=('공헌이익개선', '개선과제', '효율화', '원가절감')):
    """기준 자료(연결패키지)의 과제 시트 → [{'연': 2026, '월': 9, '구분':…, '항목':…, '내용':…,
    '금액': float|None, '상태':…, '담당':…}, …]. 시트가 없으면 None, 비어 있으면 []."""
    시트 = next((n for n in 시트들 if n in xls.sheet_names), None)
    if 시트 is None:
        return None
    d = pd.read_excel(xls, sheet_name=시트, header=None, dtype=object)
    머리 = next((i for i in range(min(len(d), 30)) if any(str(v).strip() == '작성월' for v in d.iloc[i].tolist())), None)
    if 머리 is None:
        return _효율화블록읽기(d)
    열 = {str(v).strip(): c for c, v in enumerate(d.iloc[머리].tolist()) if isinstance(v, str)}
    def 칸(i, 이름):
        c = next((c for k, c in 열.items() if k.startswith(이름)), None)
        v = d.iat[i, c] if c is not None else None
        return None if v is None or (isinstance(v, float) and pd.isna(v)) else v
    out = []
    for i in range(머리 + 1, len(d)):
        월값 = 칸(i, '작성월')
        내용 = 칸(i, '진행 내용')
        항목 = 칸(i, '개선과제') or 칸(i, '효율화 항목') or 칸(i, '절감 항목')
        if 월값 is None or (내용 is None and 항목 is None):
            continue
        연 = 월 = None
        if hasattr(월값, 'year'):
            연, 월 = int(월값.year), int(월값.month)
        else:
            t = str(월값).strip()
            m = re.match(r'^(20\d\d)\s*[-./년]\s*(\d{1,2})', t) or re.match(r'^(20\d\d)(\d\d)$', t)
            if m:
                연, 월 = int(m.group(1)), int(m.group(2))
            else:
                m = re.match(r'^(\d{1,2})\s*월?$', t)
                if m:
                    월 = int(m.group(1))
        if not 월 or not 1 <= 월 <= 12:
            continue
        금액 = pd.to_numeric(pd.Series([칸(i, '월 개선') if 칸(i, '월 개선') is not None
                                       else 칸(i, '절감 금액')]), errors='coerce').iat[0]
        out.append({'연': 연, '월': 월, '구분': str(칸(i, '구분') or ''), '항목': str(항목 or ''),
                    '내용': str(내용 or ''), '금액': None if pd.isna(금액) else float(금액),
                    '상태': str(칸(i, '진행 상태') or ''), '담당': str(칸(i, '담당') or '')})
    return out


def _패키지BS밑자료(xls):
    """연결패키지 BS·IS 시트(E 국문 과목 · G 전기 · I 당기 수정전)에서 과목별 금액을 읽습니다."""
    out = {}
    for 시트, 첫줄 in (('BS', 8), ('IS', 7)):
        d = pd.read_excel(xls, sheet_name=시트, header=None)
        전기, 당기 = {}, {}
        for i in range(첫줄, len(d)):
            국 = d.iat[i, 4] if d.shape[1] > 4 else None
            if 국 is None or (isinstance(국, float) and pd.isna(국)):
                continue
            국 = str(국).strip()
            g = pd.to_numeric(pd.Series([d.iat[i, 6]]), errors='coerce').iat[0]
            v = pd.to_numeric(pd.Series([d.iat[i, 8]]), errors='coerce').iat[0]
            if pd.isna(g) and pd.isna(v):
                continue
            전기[국] = 전기.get(국, 0.0) + (0.0 if pd.isna(g) else float(g))
            당기[국] = 당기.get(국, 0.0) + (0.0 if pd.isna(v) else float(v))
        out[시트 + '전기'], out[시트 + '당기'] = 전기, 당기
    return out


glob_소계 = re.compile(r'^\s*Total\b', re.I)
# 자산 자리에 적혀 있지만 실제로는 부채라 부호를 한 번 더 뒤집는 계정과목
BS자산칸부채 = {'부가세대급금'}


def _BS밑자료(xls, 매핑):
    """연결패키지 BS 에 넣을 「국문 계정과목별 금액」을 만듭니다.

    전기 : 26년 BS기초 시트 (F열 CTK 연결 계정과목 · B열 최종확정 금액)
    당기 : 26년 원장_raw 에서 계정별로 기초 + 증감
           ※ 자산은 기초+(차변−대변), 부채·자본은 기초−(차변−대변) —
             퀵북 재무상태표가 부채·자본을 양수로 보여 주는 방식에 맞춥니다.
    """
    전기 = {}
    if '26년 BS기초' in xls.sheet_names:
        d = pd.read_excel(xls, sheet_name='26년 BS기초', header=None)
        for _i, 줄 in d.iterrows():
            # ※ 시트의 「CTK 연결 계정과목」 칸이 아니라, 당기와 똑같이 BS_IS_매핑을 태웁니다.
            #   (그 칸은 매출채권을 미수금으로 적어 두어 당기와 이름이 어긋납니다)
            계정 = 줄.iloc[0]
            if 계정 is None or (isinstance(계정, float) and pd.isna(계정)):
                continue
            글 = str(계정).strip()
            if glob_소계.match(글):      # 「Total …」 줄은 밑줄과 겹쳐 두 번 세집니다
                continue
            키 = 글.lower()
            과목 = str(매핑['kor'].get(키, '')).strip()
            if not 과목 or 매핑['bsis'].get(키) != 'BS':
                continue
            v = pd.to_numeric(pd.Series([줄.iloc[1]]), errors='coerce').fillna(0).iat[0]
            전기[과목] = 전기.get(과목, 0.0) + float(v)

    raw = pd.read_excel(xls, sheet_name='26년 원장_raw', header=4)
    계정칸 = _글자칸(raw, raw.columns[0])
    그룹, 지금 = [], ''
    for v in 계정칸:                      # 계정 이름은 A열 머리줄에 한 번만 적혀 있습니다
        if v:
            지금 = v
        그룹.append(지금)
    이름 = _글자칸(raw, 'Distribution account')
    차 = pd.to_numeric(raw.get('Debit'), errors='coerce').fillna(0)
    대 = pd.to_numeric(raw.get('Credit'), errors='coerce').fillna(0)
    잔 = pd.to_numeric(raw.get('Balance'), errors='coerce').fillna(0)
    날 = pd.to_datetime(raw['Transaction date'], errors='coerce')
    기초, 증감 = {}, {}
    for i in range(len(raw)):
        if 이름[i] == 'Beginning Balance':
            기초[그룹[i]] = 기초.get(그룹[i], 0.0) + float(잔.iat[i])
        elif 이름[i] and pd.notna(날.iat[i]):
            증감[이름[i]] = 증감.get(이름[i], 0.0) + float(차.iat[i]) - float(대.iat[i])

    한글 = {k: v for k, v in 매핑['kor'].items()}
    비에스 = 매핑['bsis']
    당기 = {}
    for 계정 in set(기초) | set(증감):
        키 = 계정.strip().lower()
        과목 = str(한글.get(키, '')).strip()
        if not 과목 or 비에스.get(키) != 'BS':
            continue
        b, m = 기초.get(계정, 0.0), 증감.get(계정, 0.0)
        # 부채·자본은 대변이 늘어나는 쪽이라 부호를 뒤집어 양수로 보여 줍니다
        # (어느 쪽인지는 BS 서식에서 「부채」 머리줄 아래에 있는지로 가립니다)
        if 과목 in BS자산칸부채:      # 자산 자리에 있지만 실제로는 부채인 계정
            값 = m - b
        elif 과목 in BS부채자본:
            값 = b - m
        else:
            값 = b + m
        당기[과목] = 당기.get(과목, 0.0) + 값
    return {'전기': 전기, '당기': 당기}


def 손익표(df, 기준='손익', 상위=7):
    """월별 손익계산서(1~12월). 비용은 양수로 표시.

    기준='손익' : 판관비를 손익계산서 계정과목으로 나눔 (금액 큰 순 상위 N개 + 기타)
    기준='활동' : 판관비를 활동분류(보고항목) 14개로 나눔
    어느 쪽이든 매출액·영업이익·EBITDA 같은 합계 줄은 똑같습니다.
    """
    d = df[df['계정분류'].eq('IS')]
    수준 = {}                      # 줄마다 들여쓰기 단계 (0 굵게 / 1 / 2)

    def S(mask):
        return d[mask].groupby('월')['금액'].sum().reindex(range(1, 13), fill_value=0.0)

    o = {}
    o['매출액'] = S(d['분류'].eq('매출'))
    o['상품매출'] = S(d['계정과목'].isin(['제품매출', '상품매출']))
    o['용역매출 (중개 · IOR)'] = S(d['계정과목'].isin(['기타매출', '용역수익', '용역매출']))
    o['매출원가'] = -S(d['분류'].eq('매출원가'))
    o['매출총이익'] = o['매출액'] - o['매출원가']
    판관 = d['분류'].eq('판관비')
    o['판매관리비'] = -S(판관)
    리스계정 = d['계정영문'].astype(str).str.strip().str.lower().eq('depreciation - rou - lease')

    def 상각쪼개기(이름, 마스크, 부모유지):
        """EBITDA 에서 무엇을 더하고 뺐는지 보이도록 상각비를 둘로 나눕니다.

        부모유지=True (활동 기준) : 「상각비」 아래에 감가상각비 · 리스상각비를 답니다.
        부모유지=False (손익 기준): 부모 이름이 「감가상각비」와 겹치므로,
                                    부모 자리를 감가상각비(리스 제외분)로 바꾸고
                                    리스상각비를 같은 단계로 나란히 놓습니다.
        """
        리스분 = -S(마스크 & 리스계정)
        본체 = o[이름] - 리스분
        깊이 = 2 if 부모유지 else 1
        if 부모유지:
            o['감가상각비'] = 본체
        else:
            o[이름] = 본체                 # 부모 줄을 그대로 재사용 (이름 중복 방지)
            수준[이름] = 1
        o['리스상각비'] = 리스분
        수준.setdefault('감가상각비', 깊이)
        수준['리스상각비'] = 깊이

    if 기준 == '활동':
        for it in 판관비항목:
            o[it] = -S(d['보고항목'].eq(it))
            수준[it] = 1
            if it == '상각비':
                상각쪼개기(it, d['보고항목'].eq(it), 부모유지=True)
    else:
        큰순 = (-d[판관].groupby('계정과목')['금액'].sum()).sort_values(ascending=False)
        고른계정 = [str(a) for a in 큰순.index[:상위] if str(a) != 'nan']
        for acc in 고른계정:
            마스크 = 판관 & d['계정과목'].eq(acc)
            o[acc] = -S(마스크)
            수준[acc] = 1
            if acc in ('감가상각비', '무형자산상각비'):
                상각쪼개기(acc, 마스크, 부모유지=False)
        # 「기타」는 상위 계정을 뺀 나머지 — 리스상각비는 상위 계정에서 갈라져 나온 것이라 함께 뺍니다
        고른합 = sum((o[a] for a in 고른계정), pd.Series(0.0, index=range(1, 13)))
        if '리스상각비' in o:
            고른합 = 고른합 + o['리스상각비']
        o['기타'] = o['판매관리비'] - 고른합
        수준['기타'] = 1
    # ★대시보드와 같은 기준★ — 판관비에서 SR&ED 를 빼고, EBITDA 에서 리스상각은 제외
    sred = (df[df['계정영문'].astype(str).str.strip().str.lower().eq('sr&ed claim')]
            .groupby('월')['금액'].sum().reindex(range(1, 13), fill_value=0.0))
    o['SR&ED 차감'] = -sred
    o['판관비 (SR&ED 반영)'] = o['판매관리비'] - sred
    o['영업이익(손실)'] = o['매출총이익'] - o['판관비 (SR&ED 반영)']
    o['영업외손익'] = S(d['분류'].eq('영업외손익'))
    o['세전이익'] = o['영업이익(손실)'] + o['영업외손익']
    상각계정 = d['계정과목'].isin(['감가상각비', '무형자산상각비'])
    상각 = -S(상각계정)
    리스상각 = -S(상각계정 & 리스계정)
    o['EBITDA 이익(손실)'] = o['영업이익(손실)'] + 상각 - 리스상각
    t = pd.DataFrame(o).T
    t.columns = list(range(1, 13))
    기본수준 = {'매출액': 0, '상품매출': 1, '용역매출 (중개 · IOR)': 1, '매출원가': 0, '매출총이익': 0,
                '판매관리비': 0, 'SR&ED 차감': 1, '판관비 (SR&ED 반영)': 0,
                '영업이익(손실)': 0, '영업외손익': 0, '세전이익': 0, 'EBITDA 이익(손실)': 0}
    t.attrs['수준'] = {**기본수준, **수준}
    return t


# 줄 들여쓰기 단계는 손익표()가 표마다 t.attrs['수준'] 으로 함께 돌려줍니다.

# 실적집계 두 화면에서는 빼고 보여줄 줄 (다른 화면 계산에는 계속 씁니다)
집계제외 = ['영업외손익', '세전이익']

# 대시보드 화면에서는 EBITDA 줄도 빼고 보여줍니다 (2026-09-23 사용자 요청).
# 엑셀에는 그대로 남깁니다 — 계산은 계속 하고 있으니 나중에 다시 넣을 수 있습니다.
화면제외 = 집계제외 + ['EBITDA 이익(손실)']

# 오른쪽 세부내역에서 고를 수 없게 할 줄 — 합계·계산으로 만들어진 줄입니다
드릴제외 = ('매출총이익', '영업이익(손실)', 'EBITDA 이익(손실)', 'SR&ED 차감',
            '판관비 (SR&ED 반영)')


def 최하위줄(줄들, 수준표, 제외=()):
    """표에서 「더 이상 쪼개지지 않는」 줄만 골라 냅니다.

    바로 다음 줄이 더 깊이 들여쓰기되어 있으면 그 줄은 하위 항목을 거느린
    상위 줄이므로 (매출액 · 판매관리비 · 감가상각비처럼) 오른쪽 세부내역
    목록에서는 빼고, 하위 줄만 남깁니다. 하위가 없는 줄은 그대로 씁니다.
    """
    골라낸 = []
    for i, 이름 in enumerate(줄들):
        내수준 = 수준표.get(이름, 1)
        다음수준 = 수준표.get(줄들[i + 1], 1) if i + 1 < len(줄들) else -1
        if 다음수준 > 내수준 or 이름 in 제외:
            continue
        골라낸.append(이름)
    return 골라낸
# ══════════════════════════════════════════════════════════════
# 2-2. 월간실적보고용 공통 계산 (통화 · SR&ED · 현금)
# ══════════════════════════════════════════════════════════════
def SRED월별(df):
    """원장의 SR&ED claim 계정 금액(월별). 재무상태표 계정이라 손익에는 안 잡힙니다."""
    m = df['계정영문'].astype(str).str.strip().str.lower().eq('sr&ed claim')
    return df[m].groupby('월')['금액'].sum().reindex(range(1, 13), fill_value=0.0)


def 현금잔고월별(df, 기초, 매핑=None):
    매핑 = 매핑 or 매핑표      # 평소엔 화면이 읽어 둔 매핑표, 엑셀작성 화면에선 넘겨받은 것
    현금 = {k for k, v in 매핑['kor'].items()
            if 매핑['bsis'].get(k) == 'BS' and isinstance(v, str)
            and ('현금' in v or '예금' in v)}
    m = df['계정영문'].astype(str).str.strip().str.lower().isin(현금)
    증감 = (-df[m].groupby('월')['금액'].sum()).reindex(range(1, 13), fill_value=0.0)
    return 기초 + 증감.cumsum(), 증감


# ── 자금 변동(현금흐름) 분류 규칙 — 필요하면 여기만 고치면 됩니다 ─────
자금분류 = {
    'AR':    {'매출채권', '미수금', '선수금'},
    'AP':    {'매입채무', '미지급금', '미지급비용', '선급금'},   # CTKonnect : 선급금 = 중개 매입대금 선지급 (2026-10-29)
    'PAY':   {'예수금', '예수금(단체보험)', '급여', '복리후생비', '퇴직급여', '사외적립자산'},
    'EQ':    {'보통주자본금', '주식발행초과금', '자본잉여금', '기타자본'},
    'DEBT':  {'단기차입금', '장기차입금', '전환사채', '단기대여금', '장기대여금'},
    'LEASE': {'리스부채(유동)', '리스부채(비유동)'},
    'CAPEX': {'기계장치', '건설중인자산', '시설장치', '비품', '공구와기구',
              '무형자산', '특허권', '사용권자산', '사용권자산(비유동)'},
    'INV':   {'기타금융자산(유동)', '기타금융자산(비유동)', '관계기업투자자산', '종속기업투자자산',
              '단기금융상품', '장기금융상품', '매도가능증권', '장기투자증권'},     # 2026-11-03 : 금융자산 투자 (씨티케이커넥트 선물거래예치금 · 매도가능증권)
    'FX':    {'외환차손익', '외화환산손익'},
}
# SR&ED 환급금이 들어온 것은 상대 계정이 매출채권이라도 영업 회수가 아니라
# 「기타입금」으로 따로 봅니다.
# ※ 반드시 「현금 계정에 돈이 들어온 줄 자체」의 적요·전표번호만 봅니다.
#   전표 전체를 보면 감가상각 전표(SR&ED MFG equipment)나 SR&ED 대출이자까지
#   덩달아 끌려 들어옵니다.
기타입금표시 = re.compile(r'sr\s*&?\s*ed|sred', re.I)


def _기타입금줄(df, 현금행, 차변):
    """현금이 들어온 줄 가운데 SR&ED 환급인 것을 가려냅니다."""
    글 = pd.Series('', index=df.index)
    for c in ('Description', '#'):
        if c in df.columns:
            글 = 글 + ' ' + df[c].fillna('').astype(str)
    들어옴 = pd.to_numeric(차변, errors='coerce').fillna(0) > 0
    return 현금행 & 들어옴 & 글.str.contains(기타입금표시, regex=True, na=False)


채권구간 = ['within Due', '~30', '~60', '~90', '~120', '121~']


def _만기일(거래일):
    """만기 = 매출한 달의 「다음 달 말일」."""
    해, 달 = 거래일.year + (거래일.month == 12), 거래일.month % 12 + 1
    return pd.Timestamp(해, 달, calendar.monthrange(해, 달)[1])


def _채권구간(경과일):
    if 경과일 <= 0:
        return 채권구간[0]
    return (채권구간[1] if 경과일 <= 30 else 채권구간[2] if 경과일 <= 60
            else 채권구간[3] if 경과일 <= 90 else 채권구간[4] if 경과일 <= 120
            else 채권구간[5])


def 채권나이표(ar, 기초채권, 기준일, 구분맵=None):
    """거래처별 매출채권을 선입선출로 남겨 경과 구간에 나눕니다.

      · 만기 = 매출한 달의 다음 달 말일 (예: 6월 매출 → 7월 31일)
      · 회수는 오래된 매출부터 차례로 지웁니다 (선입선출)
      · 그 거래처의 올해 매출보다 더 받은 돈은 지난해에서 넘어온
        「기초채권」을 갚은 것으로 보고 기초채권에서 뺍니다
    """
    d = ar.sort_values(['거래일', 'TransactionID'], kind='stable')
    이름 = d['Name'].apply(lambda v: '' if v is None or (isinstance(v, float) and pd.isna(v))
                           else re.sub(r'\s*\((IC\s+)?(Customer|Customers|Vendor)\)\s*$', '', str(v).strip()))
    _대표 = {}
    for n in 이름:
        if n:
            _대표.setdefault(n.upper(), n)
    이름 = 이름.map(lambda n: _대표.get(n.upper(), n))
    차 = pd.to_numeric(d.get('Debit'), errors='coerce').fillna(0).tolist()
    대 = pd.to_numeric(d.get('Credit'), errors='coerce').fillna(0).tolist()
    날 = list(d['거래일'])
    이름 = 이름.tolist()
    전표 = list(d['TransactionID']) if 'TransactionID' in d.columns else [None] * len(d)
    구분맵 = 구분맵 or {}

    남은, 초과 = {}, 0.0
    초과거래처 = {}
    for i in range(len(d)):
        n = 이름[i] or '(거래처 미상)'
        큐 = 남은.setdefault(n, [])
        돈 = 0.0
        if 차[i] > 0:
            큐.append([날[i], float(차[i]), 구분맵.get(전표[i], '')])
            돈 = 초과거래처.pop(n, 0.0)          # 먼저 더 받아 둔 돈(선수)은 새 청구에 먼저 씀
            초과 -= 돈
        돈 += float(대[i])
        while 돈 > 1e-9 and 큐:
            뺄 = min(돈, 큐[0][1])
            큐[0][1] -= 뺄
            돈 -= 뺄
            if 큐[0][1] <= 1e-9:
                큐.pop(0)
        초과 += 돈                      # 올해 매출을 다 갚고도 남은 돈 → 기초채권에서
        if 돈 > 1e-9:
            초과거래처[n] = 초과거래처.get(n, 0.0) + 돈

    줄 = []
    for n, 큐 in 남은.items():
        묶음 = {}
        for 날짜, 금액, 구 in 큐:
            if 금액 <= 0.005:
                continue
            칸 = 묶음.setdefault(구, {c: 0.0 for c in 채권구간})
            칸[_채권구간((기준일 - _만기일(날짜)).days)] += 금액
        for 구, 칸 in 묶음.items():
            if sum(칸.values()) > 0.5:
                줄.append({'매출구분': 구, '거래처명': n, **칸, '채권총액': sum(칸.values())})

    기초남음 = max(float(기초채권) - 초과, 0.0)
    if 기초남음 > 0.5:                  # 25년에서 넘어온 몫 — 만기가 26년 1월말이라 121일↑
        칸 = {c: 0.0 for c in 채권구간}
        칸['121~'] = 기초남음
        줄.append({'매출구분': '', '거래처명': '기초채권 (거래처 없음)', **칸, '채권총액': 기초남음})

    표 = pd.DataFrame(줄, columns=['매출구분', '거래처명'] + 채권구간 + ['채권총액'])
    if len(표):
        표 = 표.sort_values('채권총액', ascending=False).reset_index(drop=True)
    표.attrs['초과거래처'] = {k: v for k, v in 초과거래처.items() if v > 0.5}
    return 표, 초과


def 자금변동(df, 매핑=None):
    """기초 현금에서 출발해 현금이 무엇 때문에 늘고 줄었는지 월별로 가릅니다.

    원장은 복식부기라, 현금 계정이 들어간 전표(TransactionID)에서 현금이 아닌
    상대 계정의 (대변−차변)이 곧 그 성격의 현금 증감이 됩니다. 은행↔은행 이체처럼
    현금끼리 오간 전표는 상대 계정이 없어 자동으로 상계됩니다.
    """
    매핑 = 매핑 or 매핑표      # 평소엔 화면이 읽어 둔 매핑표, 엑셀작성 화면에선 넘겨받은 것
    현금 = {k for k, v in 매핑['kor'].items()
            if 매핑['bsis'].get(k) == 'BS' and isinstance(v, str)
            and ('현금' in v or '예금' in v)}
    키 = df['계정영문'].astype(str).str.strip().str.lower()
    현금행 = 키.isin(현금)
    if 'TransactionID' not in df.columns:
        return None, None
    전표 = df['TransactionID']
    현금전표 = set(전표[현금행].dropna())
    상대 = df[전표.isin(현금전표) & ~현금행].copy()
    if 상대.empty:
        return None, None
    상대['영향'] = 상대['Debit'].mul(-1).add(상대['Credit'])      # 대변−차변 = 현금 증감

    이름 = 상대['계정과목'].astype(str)

    def 구분(n):
        for k, s in 자금분류.items():
            if n in s:
                return k
        return 'OTHOP'
    상대['구분'] = [구분(n) for n in 이름]
    상대.loc[이름.str.contains('감가상각누계', na=False), '구분'] = 'CAPEX'
    # SR&ED 환급금이 들어온 전표는 매출채권 회수가 아니라 기타입금
    같은전표 = set(전표[_기타입금줄(df, 현금행, df.get('Debit'))].dropna())
    상대.loc[상대['TransactionID'].isin(같은전표), '구분'] = 'ETCIN'

    월들 = list(range(1, 13))
    표 = (상대.pivot_table(index='구분', columns='월', values='영향', aggfunc='sum')
          .reindex(columns=월들).fillna(0.0))

    def 줄(k):
        return 표.loc[k] if k in 표.index else pd.Series(0.0, index=월들)

    def 양(k):
        return 줄(k).clip(lower=0)

    def 음(k):
        return 줄(k).clip(upper=0)

    o = {}
    o['AR'], o['AP'], o['PAY'], o['OTHOP'] = 줄('AR'), 줄('AP'), 줄('PAY'), 줄('OTHOP')
    o['ETCIN'] = 줄('ETCIN')          # 정부지원금 환급 등 매출채권 회수가 아닌 입금
    o['지급계'] = o['AP'] + o['PAY'] + o['OTHOP']
    o['영업'] = o['AR'] + o['ETCIN'] + o['지급계']
    o['CAPEX'] = 줄('CAPEX')
    o['INV'] = 줄('INV')
    o['투자'] = o['CAPEX'] + o['INV']
    o['EQ'] = 줄('EQ')
    # 차입은 한 줄로 봅니다 — 빌리면 +, 갚으면 − (같은 달에 둘 다 있으면 서로 상계)
    o['차입'] = 양('DEBT')
    o['상환'] = 음('DEBT')
    o['차입증감'] = 줄('DEBT')
    o['리스'] = 줄('LEASE')
    o['재무'] = o['EQ'] + o['차입증감'] + o['리스']
    o['FX'] = 줄('FX')
    o['순증감'] = o['영업'] + o['투자'] + o['재무'] + o['FX']
    t = pd.DataFrame(o).T
    t.columns = 월들
    return t, 상대[['구분', '영향', '월', 'Name', '계정과목']]






# ══════════════════════════════════════════════════════════════
# 분석 엑셀 만들기 — 올린 원장 하나로 보고용 엑셀을 즉석에서 생성
# ══════════════════════════════════════════════════════════════
def 분석엑셀만들기(원장, 전년, 손익표본, 활동표본, 자금, 잔고, 기초현금액, 연도, 월끝, 매핑,
                   원천시트=None):
    """대시보드가 보여 주는 표를 그대로 담은 엑셀을 만들어 bytes 로 돌려줍니다.

    ★피벗·외부링크를 쓰지 않습니다★ — 엑셀에서 「복구됨」이 뜰 여지를 없앴습니다.
    「원장_처리본」 시트가 원천이고, 손익 표는 그 시트를 SUMIFS 로 집계하므로
    원장을 고치면 엑셀 안에서 숫자가 다시 계산됩니다.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    NAVY, OR_, GREY, RED = 'FF12294D', 'FFE2611C', 'FF8A8880', 'FFD03B3B'
    글꼴 = '맑은 고딕'
    머리채움 = PatternFill('solid', fgColor=NAVY)
    띠채움 = PatternFill('solid', fgColor='FFF7F8FA')
    누적채움 = PatternFill('solid', fgColor='FFF1F4F9')
    가는 = Side(style='thin', color='FFE7E7E1')
    테두리 = Border(left=가는, right=가는, top=가는, bottom=가는)
    돈 = '#,##0;(#,##0);-'
    월들 = list(range(1, 월끝 + 1))

    wb = Workbook()

    def 표머리(ws, 제목, 부제, 열들, 시작행=1):
        c = ws.cell(시작행, 1, 제목)
        c.font = Font(글꼴, size=15, bold=True, color=NAVY)
        ws.cell(시작행 + 1, 1, 부제).font = Font(글꼴, size=9, color=GREY)
        hr = 시작행 + 3
        for i, (이름, 폭) in enumerate(열들, 1):
            h = ws.cell(hr, i, 이름)
            h.font = Font(글꼴, size=10, bold=True, color='FFFFFFFF')
            h.fill = 머리채움
            h.alignment = Alignment('center', 'center', wrap_text=True)
            h.border = 테두리
            ws.column_dimensions[get_column_letter(i)].width = 폭
        ws.freeze_panes = ws.cell(hr + 1, 2)
        return hr + 1

    # ── 1. 읽어보기 ────────────────────────────────────────
    ws = wb.active
    ws.title = '읽어보기'
    ws.sheet_view.showGridLines = False
    for col, w in zip('AB', (26, 104)):
        ws.column_dimensions[col].width = w
    ws['A1'] = f'CTKonnect, Inc. {연도}년 1~{월끝}월 실적분석'
    ws['A1'].font = Font(글꼴, size=17, bold=True, color=NAVY)
    ws['A2'] = '이 파일은 대시보드가 원장에서 자동으로 계산해 만든 것입니다 (클로드 작성)'
    ws['A2'].font = Font(글꼴, size=10, color=OR_, bold=True)
    안내 = [
        ('만든 날', datetime.datetime.now().strftime('%Y-%m-%d %H:%M')),
        ('만든 주체', '클로드(Claude) — CTKonnect 실적보고 대시보드'),
        ('원천 자료', f'업로드한 원장 {len(원장):,}건 (당해 {연도}년) · 전년 {len(전년):,}건'),
        ('기준 기간', f'{연도}년 1월 ~ {월끝}월'),
        ('통화', '전부 USD (환산 없음)'),
        ('', ''),
        ('■ 시트 구성', ''),
        ('원장_처리본', '업로드한 원장에 계정분류·계정과목·월을 붙인 것. 아래 표들의 원천입니다'),
        ('월별 손익계산서', '손익계산서 계정과목 기준. 원장_처리본을 SUMIFS 로 집계 — 원장을 고치면 다시 계산됩니다'),
        ('월별 손익(활동기준)', '활동분류 기준으로 본 같은 손익'),
        ('당월 실적집계', f'{월끝}월을 전월·전년동월·누적과 나란히 비교'),
        ('현금흐름', '기초 현금에서 출발해 월별로 무엇 때문에 늘고 줄었는지'),
        ('미수채권', '거래처별 매출채권 잔액'),
    ] + ([
        ('○○년 원장_raw', '원장 원본 + 활동분류. ★대시보드가 읽는 시트라 지우면 안 됩니다★'),
        ('BS_IS_매핑', '영문 계정 → 한글 계정과목·BS/IS 구분. 이 시트가 있어야 대시보드가 그려집니다'),
        ('월간보고_입력', '월간실적 보고서에 손으로 넣는 값'),
        ('◆ 다시 쓰는 법', '이 파일을 대시보드 왼쪽 「⚙ 설정 · 파일 업로드」에 그대로 올리면 화면이 그려집니다. '
                           '다음 달에는 이 파일을 「전월 실적보고자료」로 올려 주세요'),
    ] if 원천시트 else []) + [
        ('', ''),
        ('■ 계산 기준', ''),
        ('영업이익', '매출총이익 − 판관비(SR&ED 차감 후)'),
        ('현금흐름', '현금 계정이 든 전표에서 현금이 아닌 상대 계정의 (대변−차변)을 그 성격의 현금 증감으로 봅니다'),
        ('', ''),
        ('■ 꼭 알아두실 점', ''),
        ('활동분류', '판관비의 활동분류(대분류·활동세부)는 원장에 이미 적혀 있는 것을 그대로 씁니다. '
                     '원장에 없으면 「기타관리경비」로 잡히므로, 새 거래는 원장에서 먼저 분류해 주셔야 합니다'),
        ('피벗 없음', '이 파일에는 피벗테이블과 외부 파일 링크가 없습니다 — 「복구됨」이 뜨지 않습니다'),
        ('25년 자료', '전년 비교는 업로드한 파일의 25년 원장을 그대로 씁니다'),
    ]
    r = 4
    for 이름, 값 in 안내:
        if 이름.startswith('■'):
            ws.cell(r, 1, 이름).font = Font(글꼴, size=11, bold=True, color=NAVY)
        elif 이름:
            ws.cell(r, 1, 이름).font = Font(글꼴, size=10, bold=True)
            c = ws.cell(r, 2, 값)
            c.font = Font(글꼴, size=10)
            c.alignment = Alignment(vertical='top', wrap_text=True)
        r += 1

    # ── 2. 원장_처리본 ─────────────────────────────────────
    원 = 원장.copy()
    원['거래일'] = pd.to_datetime(원['거래일']).dt.strftime('%Y-%m-%d')
    보낼열 = ['거래일', '계정영문', '계정과목', '계정분류', '분류', '보고항목',
              '활동분류(대분류)', '활동세부', 'Name', 'Description',
              'Debit', 'Credit', '금액', '월']
    보낼열 = [c for c in 보낼열 if c in 원.columns]
    ws2 = wb.create_sheet('원장_처리본')
    ws2.sheet_view.showGridLines = False
    폭 = {'거래일': 12, '계정영문': 30, '계정과목': 16, '계정분류': 9, '분류': 10,
          '보고항목': 15, '활동분류(대분류)': 18, '활동세부': 20, 'Name': 28,
          'Description': 40, 'Debit': 14, 'Credit': 14, '금액': 14, '월': 6}
    hr2 = 표머리(ws2, '원장 처리본',
                 f'업로드한 원장에 계정분류·계정과목·월을 붙였습니다 · {len(원):,}건 · 아래 표들의 원천',
                 [(c, 폭.get(c, 14)) for c in 보낼열])
    for i, 행 in enumerate(원[보낼열].itertuples(index=False), hr2):
        for j, v in enumerate(행, 1):
            c = ws2.cell(i, j, None if (isinstance(v, float) and pd.isna(v)) else v)
            c.font = Font(글꼴, size=9)
            if 보낼열[j - 1] in ('Debit', 'Credit', '금액'):
                c.number_format = '#,##0.00'
    끝행 = hr2 + len(원) - 1
    열자리 = {c: get_column_letter(i + 1) for i, c in enumerate(보낼열)}
    범위 = lambda c: f"'원장_처리본'!${열자리[c]}${hr2}:${열자리[c]}${끝행}"

    # ── 3. 월별 손익계산서 (SUMIFS 로 다시 계산됨) ──────────
    def 손익시트(제목, 표, 기준열):
        ws3 = wb.create_sheet(제목)
        ws3.sheet_view.showGridLines = False
        열들 = [('구분', 30)] + [(f'{m}월', 14) for m in 월들] + [('누적', 15)]
        hr3 = 표머리(ws3, 제목,
                     f'{연도}년 · 단위 USD · 비용은 양수 · 원장_처리본을 SUMIFS 로 집계합니다', 열들)
        수준 = 표.attrs['수준']
        r3 = hr3
        for 이름 in 표.index:
            if 이름 in 집계제외:
                continue
            lv = 수준.get(이름, 1)
            c = ws3.cell(r3, 1, ('    ' * lv) + 이름)
            c.font = Font(글꼴, size=10, bold=(lv == 0))
            c.border = 테두리
            if lv == 0:
                c.fill = 띠채움
            계정식 = None
            if 기준열 and lv >= 1 and 이름 in set(원[기준열].dropna().astype(str)):
                계정식 = 이름
            for k, m in enumerate(월들):
                cc = ws3.cell(r3, k + 2)
                if 계정식:
                    cc.value = (f'=-SUMIFS({범위("금액")},{범위(기준열)},$A{r3}&"",'
                                f'{범위("월")},{get_column_letter(k + 2)}${hr3 - 1})')
                else:
                    cc.value = float(표.loc[이름, m])
                cc.number_format = 돈
                cc.font = Font(글꼴, size=10, bold=(lv == 0),
                               color=(RED if float(표.loc[이름, m]) < 0 else 'FF000000'))
                cc.border = 테두리
                if lv == 0:
                    cc.fill = 띠채움
            L = get_column_letter(2)
            R = get_column_letter(len(월들) + 1)
            t = ws3.cell(r3, len(월들) + 2, f'=SUM({L}{r3}:{R}{r3})')
            t.number_format = 돈
            t.font = Font(글꼴, size=10, bold=True)
            t.fill = 누적채움
            t.border = 테두리
            r3 += 1
        # 머리글 줄에 월 숫자를 숨겨 두어 SUMIFS 가 참조하게 함
        for k, m in enumerate(월들):
            h = ws3.cell(hr3 - 1, k + 2, m)
            h.font = Font(글꼴, size=10, bold=True, color='FFFFFFFF')
            h.fill = 머리채움
            h.alignment = Alignment('center', 'center')
            h.border = 테두리
        return ws3

    손익시트('월별 손익계산서', 손익표본, '계정과목')
    손익시트('월별 손익(활동기준)', 활동표본, '보고항목')

    # ── 4. 당월 실적집계 ───────────────────────────────────
    ws4 = wb.create_sheet('당월 실적집계')
    ws4.sheet_view.showGridLines = False
    이전 = 월끝 - 1 if 월끝 > 1 else None
    열들 = [('구분', 30), (f'당월({월끝}월)', 15), (f'전월({이전}월)' if 이전 else '전월', 15),
            ('전월 증감', 15), (f'전년동월({str(연도-1)[2:]}년 {월끝}월)', 17), ('전년동월 증감', 15),
            (f'당해누적(1~{월끝}월)', 17), (f'전년동기(1~{월끝}월)', 17), ('누적 증감', 15)]
    hr4 = 표머리(ws4, f'{연도}년 {월끝}월 실적집계',
                 '전월 · 전년동월 · 누적을 나란히 비교 · 단위 USD · 비용은 양수', 열들)
    수준4 = 손익표본.attrs['수준']
    전년표 = 손익표(전년, '손익') if len(전년) else None
    r4 = hr4
    for 이름 in 손익표본.index:
        if 이름 in 집계제외:
            continue
        lv = 수준4.get(이름, 1)
        당월 = float(손익표본.loc[이름, 월끝])
        전월 = float(손익표본.loc[이름, 이전]) if 이전 else 0.0
        전동 = float(전년표.loc[이름, 월끝]) if (전년표 is not None and 이름 in 전년표.index) else 0.0
        누적 = float(손익표본.loc[이름, 월들].sum())
        전누 = float(전년표.loc[이름, 월들].sum()) if (전년표 is not None and 이름 in 전년표.index) else 0.0
        값들 = [당월, 전월, 당월 - 전월, 전동, 당월 - 전동, 누적, 전누, 누적 - 전누]
        c = ws4.cell(r4, 1, ('    ' * lv) + 이름)
        c.font = Font(글꼴, size=10, bold=(lv == 0))
        c.border = 테두리
        if lv == 0:
            c.fill = 띠채움
        for j, v in enumerate(값들, 2):
            cc = ws4.cell(r4, j, v)
            cc.number_format = 돈
            cc.font = Font(글꼴, size=10, bold=(lv == 0), color=(RED if v < 0 else 'FF000000'))
            cc.border = 테두리
            if lv == 0:
                cc.fill = 띠채움
        r4 += 1

    # ── 5. 현금흐름 ────────────────────────────────────────
    if 자금 is not None:
        ws5 = wb.create_sheet('현금흐름')
        ws5.sheet_view.showGridLines = False
        열들 = [('구분', 32)] + [(f'{m}월', 14) for m in 월들] + [('누적', 15)]
        hr5 = 표머리(ws5, '현금흐름',
                     '기초 현금에서 출발해 월별로 무엇 때문에 늘고 줄었는지 · 단위 USD', 열들)
        기초 = [float(잔고.get(m - 1, 기초현금액)) if m > 1 else float(기초현금액) for m in 월들]
        기말 = [float(잔고.get(m, 0.0)) for m in 월들]
        ㅈ = lambda k: [float(자금.loc[k, m]) for m in 월들]
        줄들 = [('기초 현금', 0, 기초, 기초[0]),
                ('영업활동', 0, ㅈ('영업'), sum(ㅈ('영업'))),
                ('    매출채권 회수', 1, ㅈ('AR'), sum(ㅈ('AR'))),
                ('    기타입금 (정부지원금 환급 등)', 1, ㅈ('ETCIN'), sum(ㅈ('ETCIN'))),
                ('    매입채무', 1, ㅈ('AP'), sum(ㅈ('AP'))),
                ('    급여', 1, ㅈ('PAY'), sum(ㅈ('PAY'))),
                ('    기타 비용', 1, ㅈ('OTHOP'), sum(ㅈ('OTHOP'))),
                ('투자활동', 0, ㅈ('투자'), sum(ㅈ('투자'))),
                ('    설비 투자', 1, ㅈ('CAPEX'), sum(ㅈ('CAPEX'))),
                ('재무활동', 0, ㅈ('재무'), sum(ㅈ('재무'))),
                ('    유상증자', 1, ㅈ('EQ'), sum(ㅈ('EQ'))),
                ('    차입증감', 1, ㅈ('차입증감'), sum(ㅈ('차입증감'))),
                ('    리스료 지급', 1, ㅈ('리스'), sum(ㅈ('리스'))),
                ('환율변동 효과', 1, ㅈ('FX'), sum(ㅈ('FX'))),
                ('순증감', 0, ㅈ('순증감'), sum(ㅈ('순증감'))),
                ('기말 현금', 0, 기말, 기말[-1])]
        r5 = hr5
        for 이름, lv, 값들, 총 in 줄들:
            c = ws5.cell(r5, 1, 이름)
            c.font = Font(글꼴, size=10, bold=(lv == 0))
            c.border = 테두리
            if lv == 0:
                c.fill = 띠채움
            for j, v in enumerate(값들 + [총], 2):
                cc = ws5.cell(r5, j, float(v))
                cc.number_format = 돈
                cc.font = Font(글꼴, size=10, bold=(lv == 0 or j == len(값들) + 2),
                               color=(RED if v < 0 else 'FF000000'))
                cc.border = 테두리
                cc.fill = 누적채움 if j == len(값들) + 2 else (띠채움 if lv == 0 else PatternFill())
            r5 += 1
        ws5.cell(r5 + 1, 1, '계산 방법 — 원장은 복식부기라, 현금이 든 전표에서 현금이 아닌 상대 계정의 '
                            '(대변−차변)이 그 돈의 성격별 현금 증감이 됩니다. 은행↔은행 이체는 서로 상계됩니다.'
                 ).font = Font(글꼴, size=9, color=GREY)

    # ── 6. 미수채권 ────────────────────────────────────────
    ar = 원장[원장['계정과목'].eq('매출채권') & 원장['Name'].notna()].copy()
    if len(ar):
        ws6 = wb.create_sheet('미수채권')
        ws6.sheet_view.showGridLines = False
        hr6 = 표머리(ws6, '거래처별 매출채권',
                     f'{연도}년 발생·회수 기준 · 단위 USD · ★기초잔액은 포함되어 있지 않습니다★',
                     [('거래처', 40), ('발생액', 16), ('회수액', 16), ('잔액', 16),
                      ('최근 거래일', 14), ('건수', 10)])
        g = ar.groupby('Name').agg(발생=('Debit', 'sum'), 회수=('Credit', 'sum'),
                                   최근=('거래일', 'max'), 건수=('Debit', 'size'))
        g['잔액'] = g['발생'] - g['회수']
        g = g.sort_values('잔액', ascending=False)
        r6 = hr6
        for 이름, x in g.iterrows():
            for j, v in enumerate([이름, x['발생'], x['회수'], x['잔액'],
                                   pd.Timestamp(x['최근']).strftime('%Y-%m-%d'), int(x['건수'])], 1):
                cc = ws6.cell(r6, j, v)
                cc.font = Font(글꼴, size=10,
                               color=(RED if (j == 4 and float(x['잔액']) < 0) else 'FF000000'))
                cc.border = 테두리
                if j in (2, 3, 4):
                    cc.number_format = '#,##0.00'
                if j in (5, 6):
                    cc.alignment = Alignment('center')
            r6 += 1

    # ── 7. 원본 시트 덧붙이기 (실적보고 엑셀작성 화면에서만 씁니다) ──
    #    이 시트들이 있어야 만든 파일을 다시 사이드바에 올렸을 때 대시보드가 그려집니다.
    if 원천시트:
        for 이름, (부제, d) in 원천시트.items():
            wsx = wb.create_sheet(이름)
            wsx.sheet_view.showGridLines = False
            머리 = 5 if 부제 is not None else 1
            if 부제 is not None:
                wsx.cell(1, 1, 'CTKonnect, Inc.').font = Font(글꼴, size=12, bold=True, color=NAVY)
                wsx.cell(2, 1, 'General Ledger List').font = Font(글꼴, size=9, color=GREY)
                wsx.cell(3, 1, 부제).font = Font(글꼴, size=9, color=GREY)
                for j, c in enumerate(d.columns, 1):
                    h = wsx.cell(머리, j, str(c))
                    h.font = Font(글꼴, size=10, bold=True, color='FFFFFFFF')
                    h.fill = 머리채움
                    h.alignment = Alignment('center', 'center')
                    wsx.column_dimensions[get_column_letter(j)].width = 16
                wsx.freeze_panes = wsx.cell(머리 + 1, 1)
            else:
                for j in range(1, min(d.shape[1], 20) + 1):
                    wsx.column_dimensions[get_column_letter(j)].width = 22
            for 행 in d.itertuples(index=False):
                wsx.append(['' if (isinstance(v, float) and pd.isna(v)) or v is pd.NaT else v
                            for v in 행])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════
# 실적보고 엑셀 만들기 — 원장 원본 + 전월 실적보고자료 → 당월 실적보고 엑셀
# ══════════════════════════════════════════════════════════════
# ★방식★ 전월 파일을 zip 그대로 열어 「○○년 원장_raw」 시트 하나만 갈아끼웁니다.
#        피벗·차트·메모·나머지 시트는 손도 대지 않으므로 그대로 살아 있고,
#        당월/월별 실적집계는 이미 SUMIFS·OFFSET 수식이라 엑셀이 열릴 때 스스로 계산합니다.
#        (파이썬으로 파일을 새로 쓰면 피벗과 차트가 사라집니다)
#
# 원장에서 저절로 나오는 것 : K 금액 · L 년 · M 월 · N 계정분류 · O 분류 · P 계정과목 ·
#                             U 보고금액 · V 월(숫자)   ← 전부 원본과 같은 수식을 채웁니다
# 전월에서 이어받는 것      : Q 활동분류 · R 활동세부 (사람이 손으로 나눈 값) · T 검토메모
엑셀기준일 = datetime.date(1899, 12, 30)
_행패턴 = re.compile(r'<row\b[^>]*\br="(\d+)"[^>]*>.*?</row>|<row\b[^>]*\br="(\d+)"[^>]*/>', re.S)
_칸패턴 = re.compile(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
_수식패턴 = re.compile(r'<f([^>]*?)>(.*?)</f>', re.S)

원장열자리 = {'A': 'Unnamed: 0', 'B': 'Distribution account', 'C': 'Transaction date',
              'D': 'Description', 'E': 'Name', 'F': 'Transaction ID', 'G': '#',
              'H': 'Balance', 'I': 'Debit', 'J': 'Credit'}
전체열 = list('ABCDEFGHIJKLMNOPQRSTUV')
자동수식열 = ['K', 'L', 'M', 'N', 'O', 'P', 'U', 'V']
이어받기열 = {'Q': '활동분류', 'R': '활동세부',
              'S': '정부지원(SR&ED/NGEN)', 'T': '클로드 검토'}


def _xml(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))


def _날짜일련(v):
    """엑셀이 쓰는 날짜 일련번호로 바꿉니다 (YEAR·MONTH 수식이 돌아가려면 숫자여야 합니다)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, str):
        p = v.strip().replace('-', '/').split('/')
        if len(p) != 3:
            return None
        try:
            d = datetime.date(int(p[2]), int(p[1]), int(p[0]))
        except Exception:
            return None
    else:
        try:
            t = pd.Timestamp(v)
            if pd.isna(t):
                return None
            d = t.date()
        except Exception:
            return None
    return (d - 엑셀기준일).days


def _이어받기키(df):
    """전월 자료와 이번 원장의 같은 거래를 짝지어 주는 「지문」.

    ※ 같은 전표·같은 금액이 여러 줄인 경우가 있어(같은 날 같은 금액의 급여 두 줄 등)
      적요(Description)까지 넣고, 그래도 겹치면 그 안에서의 순번을 붙여 갈라 냅니다.
      순번을 안 붙이면 둘 중 하나의 분류가 엉뚱한 줄에 붙습니다.
    """
    k = df['Distribution account'].fillna('∅').astype(str).str.strip()
    for c in ('Transaction ID', 'Debit', 'Credit'):
        k = k + '|' + pd.to_numeric(df.get(c), errors='coerce').round(2).fillna(-1).astype(str)
    for c in ('Name', 'Description'):
        칸 = df[c] if c in df.columns else pd.Series(index=df.index, dtype=object)
        k = k + '|' + 칸.fillna('∅').astype(str).str.strip()
    k = k + '|' + df['Transaction date'].map(lambda v: str(_날짜일련(v)))
    return k + '#' + k.groupby(k).cumcount().astype(str)


# ── 새 거래를 스스로 나눠 주는 부분 ────────────────────────────
#    전월 자료에 없던 거래는 활동분류·활동세부가 비어 있어서, 그대로 두면
#    「월별 실적집계」의 SUMIFS 가 그 줄을 못 세고 이번 달 칸이 0 으로 남습니다.
#    그래서 지난달까지 사람이 나눠 놓은 것을 배워서 새 줄에 붙이고,
#    확실하지 않은 것은 T열 「클로드 검토」에 왜 그렇게 붙였는지 적어 둡니다.
_숫자토막 = re.compile(r'\d+')
_공백토막 = re.compile(r'\s+')


def _글자정리(v):
    """적요·거래처를 견주기 좋게 다듬습니다 (대소문자·여러 칸 공백·전표번호 무시)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return _공백토막.sub(' ', _숫자토막.sub('#', str(v).lower().strip()))


def _글자칸(df, 이름):
    """한 열을 「빈칸이면 '' 인 글자 목록」으로 바꿉니다.

    ※ .astype(str) 만 쓰면 빈칸이 float nan 으로 남아 「nan」이라는 계정이 생깁니다.
    """
    if 이름 not in df.columns:
        return [''] * len(df)
    return ['' if (v is None or (isinstance(v, float) and pd.isna(v))
                   or v is pd.NA) else str(v).strip() for v in df[이름]]


def _열찾기(df, 이름):
    for c in df.columns:
        if str(c).strip() == 이름:
            return c
    for c in df.columns:
        if str(c).startswith(이름):
            return c
    return None


def _분류학습(전):
    """전월 자료에서 「이 계정·이 적요는 어떤 활동으로 나눴는지」를 배웁니다."""
    적요, 거래처, 계정, 본줄 = {}, {}, {}, collections.Counter()
    q열, r열 = _열찾기(전, '활동분류'), _열찾기(전, '활동세부')
    if q열 is None or r열 is None:
        return 적요, 거래처, 계정, 본줄
    계정칸 = _글자칸(전, 'Distribution account')
    적요칸 = [_글자정리(v) for v in _글자칸(전, 'Description')]
    이름칸 = [_글자정리(v) for v in _글자칸(전, 'Name')]
    for a, dn, nn, q, r in zip(계정칸, 적요칸, 이름칸, 전[q열], 전[r열]):
        if not a:
            continue
        본줄[a] += 1
        r = '' if (r is None or (isinstance(r, float) and pd.isna(r))) else str(r).strip()
        if not r:
            continue
        짝 = ('' if (q is None or (isinstance(q, float) and pd.isna(q)))
              else str(q).strip(), r)
        계정.setdefault(a, collections.Counter())[짝] += 1
        if dn:
            적요.setdefault((a, dn), collections.Counter())[짝] += 1
        if nn:
            거래처.setdefault((a, nn), collections.Counter())[짝] += 1
    return 적요, 거래처, 계정, 본줄


def _자동분류(새, 이어받음, 배움, 유효세부, 매핑분류):
    """이어받지 못한 줄의 Q·R 을 채우고, 그 근거를 T 에 적습니다.

    유효세부 : 「월별 실적집계」가 실제로 세는 활동세부 이름 모음
    매핑분류 : BS_IS_매핑의 계정 → 'BS' / 'IS'
    돌려주는 값 : 근거별 건수 dict
    """
    적요, 거래처, 계정, 본줄 = 배움
    계정칸 = _글자칸(새, 'Distribution account')
    적요칸 = [_글자정리(v) for v in _글자칸(새, 'Description')]
    이름칸 = [_글자정리(v) for v in _글자칸(새, 'Name')]
    날짜칸 = 새['Transaction date'].tolist()

    센것 = collections.Counter()
    메모 = 이어받음.setdefault('T', [None] * len(새))
    Q칸 = 이어받음.setdefault('Q', [None] * len(새))
    R칸 = 이어받음.setdefault('R', [None] * len(새))
    for i in range(len(새)):
        a = 계정칸[i]
        if not a:
            continue
        if _날짜일련(날짜칸[i]) is None:      # 이월잔액 · 소계 줄은 건드리지 않습니다
            continue
        아는계정 = (not 매핑분류) or (a in 매핑분류)
        if not 아는계정 and not 메모[i]:
            메모[i] = f'BS_IS_매핑에 없는 계정 「{a}」 — 계정과목·계정분류가 빈칸이 됩니다'
            센것['매핑밖'] += 1
        if R칸[i]:
            센것['이어받음'] += 1
            continue

        짝, 근거, 급함 = None, '', False
        if (a, 적요칸[i]) in 적요:
            짝, 근거 = 적요[(a, 적요칸[i])].most_common(1)[0][0], '적요일치'
        elif (a, 이름칸[i]) in 거래처:
            짝, 근거 = 거래처[(a, 이름칸[i])].most_common(1)[0][0], '거래처일치'
        elif a in 계정 and len(계정[a]) == 1:
            짝, 근거 = next(iter(계정[a])), '계정과목대표값'
        elif a in 계정:
            짝, 근거, 급함 = 계정[a].most_common(1)[0][0], '계정과목최빈값', True
        elif a in 본줄:
            센것['원래빈칸'] += 1        # 이 계정은 원래 활동분류를 안 붙이는 계정입니다
            continue
        elif 매핑분류.get(a) == 'BS':
            센것['신규BS계정'] += 1      # 재무상태표 계정이라 활동분류가 필요 없습니다
            continue
        else:
            센것['신규계정'] += 1
            if not 메모[i]:
                메모[i] = f'신규 계정 「{a}」 — 활동분류·활동세부를 새로 정해 주셔야 합니다'
            continue

        # 같은 계정인데 지난달까지 두 갈래 이상으로 나뉜 적이 있으면 더 세게 표시합니다
        if len(계정.get(a, ())) > 1:
            급함 = True
        Q칸[i] = 짝[0] or None
        R칸[i] = 짝[1]
        센것[근거] += 1

        말 = f'자동분류-{근거}({"확인필요" if 급함 else "검토요망"})'
        if 유효세부 and 짝[1] not in 유효세부:
            말 += ' ※월별 실적집계에 없는 활동세부'
            센것['집계밖'] += 1
            급함 = True
        센것['확인필요' if 급함 else '검토요망'] += 1
        if not 메모[i]:
            메모[i] = 말
    return dict(센것)


def _풀기(s):
    """XML 글자 되돌리기 — &lt; 같은 이름 참조와 &#52828; 같은 숫자 참조(openpyxl 저장 파일) 모두."""
    return _html.unescape(s)


def _공유글자(z):
    try:
        sx = z.read('xl/sharedStrings.xml').decode('utf-8')
    except KeyError:
        return []
    return [_풀기(re.sub(r'<[^>]+>', '', m.group(1)))
            for m in re.finditer(r'<si>(.*?)</si>', sx, re.S)]


def _시트글자(z, 파일, 열, 공유):
    """시트 XML 에서 한 열의 글자값만 빠르게 뽑습니다 (openpyxl 은 이 파일에 너무 느립니다)."""
    try:
        x = z.read(파일).decode('utf-8')
    except KeyError:
        return []
    if '<sheetData>' not in x:
        return []
    본문 = x.split('<sheetData>', 1)[1].rsplit('</sheetData>', 1)[0]
    값 = []
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', 본문, re.S):
        if cm.group(1) != 열:
            continue
        속성, 몸 = cm.group(2), cm.group(3) or ''
        if '<f' in 몸:
            continue
        vm = re.search(r'<v>(.*?)</v>', 몸, re.S)
        if vm:
            v = vm.group(1)
            값.append(공유[int(v)] if ' t="s"' in 속성 and v.isdigit()
                      and int(v) < len(공유) else _풀기(v))
        elif '<is>' in 몸:
            값.append(_풀기(re.sub(r'<[^>]+>', '', 몸)))
    return [str(v).strip() for v in 값 if str(v).strip()]


def _시트칸읽기(z, 파일, 열들, 공유):
    """시트 XML 에서 원하는 열들의 값을 {행번호: {열: 값}} 으로 읽습니다."""
    try:
        x = z.read(파일).decode('utf-8')
    except KeyError:
        return {}
    if '<sheetData>' not in x:
        return {}
    본문 = x.split('<sheetData>', 1)[1].rsplit('</sheetData>', 1)[0]
    표 = {}
    for rm in re.finditer(r'<row[^>]*\br="(\d+)"[^>]*>(.*?)</row>', 본문, re.S):
        행 = int(rm.group(1))
        칸 = {}
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)',
                              rm.group(2), re.S):
            열, 속성, 몸 = cm.group(1), cm.group(2), cm.group(3) or ''
            if 열 not in 열들:
                continue
            vm = re.search(r'<v>(.*?)</v>', 몸, re.S)
            if vm:
                v = vm.group(1)
                칸[열] = (공유[int(v)] if ' t="s"' in 속성 and v.isdigit()
                          and int(v) < len(공유) else _풀기(v))
            elif '<is>' in 몸:
                칸[열] = _풀기(re.sub(r'<[^>]+>', '', 몸))
        if 칸:
            표[행] = 칸
    return 표


def _매핑표읽기(z, 파일, 공유):
    """BS_IS_매핑 시트에서 「계정 이름 → (BS/IS, 한글 계정과목)」을 읽습니다."""
    지도 = {}
    for _행, 칸 in _시트칸읽기(z, 파일, ('A', 'B', 'C'), 공유).items():
        이름 = str(칸.get('A', '')).strip()
        if 이름:
            지도[이름] = (str(칸.get('B', '')).strip().upper(),
                          str(칸.get('C', '')).strip())
    return 지도


def _시트본뜨기(원본xml):
    """원본 원장 시트에서 1~5행 · 칸 서식번호 · 수식 본을 뽑아냅니다."""
    앞, 나머지 = 원본xml.split('<sheetData>', 1)
    본문, 뒤 = 나머지.rsplit('</sheetData>', 1)
    머리행, 그룹서식, 자료서식, 수식본 = {}, {}, {}, {}
    for m in _행패턴.finditer(본문):
        r = int(m.group(1) or m.group(2))
        조각 = m.group(0)
        if r <= 5:
            머리행[r] = 조각
            continue
        칸 = {}
        for cm in _칸패턴.finditer(조각):
            col, attrs, body = cm.group(1), cm.group(2), cm.group(3) or ''
            if col not in 전체열:
                continue
            s = re.search(r's="(\d+)"', attrs)
            칸[col] = (s.group(1) if s else None, body)
            fm = _수식패턴.search(body)
            if fm and col not in 수식본 and fm.group(2).strip():
                수식본[col] = (r, fm.group(2))
        표 = 그룹서식 if 칸.get('B', (None, ''))[1] in ('', None) else 자료서식
        if not 표:
            for col, (s, _b) in 칸.items():
                if s:
                    표[col] = s
        if 그룹서식 and 자료서식 and len(수식본) >= 10:
            break
    return 앞, 뒤, 머리행, 그룹서식, 자료서식, 수식본


def _수식옮기기(본, 원행, 새행):
    """C6·J6 처럼 본이 된 줄 번호를 새 줄 번호로 바꿉니다."""
    return re.sub(r'(\$?[A-Z]{1,3}\$?)' + str(원행) + r'(?![0-9])',
                  lambda m: m.group(1) + str(새행), 본)


# 「26년 원장_raw」 K~V 칸의 본디 수식 (6번째 줄 기준).
#  ※ 올려 주신 전월 파일에서 어느 칸의 수식이 지워져 있으면 여기 것으로 되살립니다.
#    (한 번 값만 남은 파일을 다시 넣으면 그 칸이 통째로 비어 버리던 것을 막습니다)
_되살린열 = []          # 이번에 되살린 칸 (화면에 알려 주려고 남깁니다)
기본수식본 = {
    'K': (6, 'IF(C6="","",J6-I6)'),
    'L': (6, 'IF(C6="","",YEAR(C6)&amp;"년")'),
    'M': (6, 'IF(C6="","",MONTH(C6)&amp;"월")'),
    'N': (6, 'IFERROR(INDEX(BS_IS_매핑!$B:$B,MATCH(B6,BS_IS_매핑!$A:$A,0)),"")'),
    'O': (6, 'IF(N6&lt;&gt;"IS","",IF(OR(P6="제품매출",P6="기타매출"),"매출",'
             'IF(P6="제품매출원가","매출원가",IF(OR(P6="이자비용",P6="잡손실",P6="잡이익",'
             'P6="외환차익",P6="외환차손",P6="외화환산이익",P6="외화환산손실",'
             'P6="유형자산처분손실"),"영업외손익","판관비"))))'),
    'P': (6, 'IFERROR(INDEX(BS_IS_매핑!$C:$C,MATCH(B6,BS_IS_매핑!$A:$A,0)),"")'),
    'Q': (6, 'IF(P6="제품매출","친환경 바이오 소재",'
             'IF(P6="기타매출","기술 서비스 및 자문",""))'),
    'R': (6, 'IF(P6="제품매출","친환경 빨대/원자재 및 부자재 공급",'
             'IF(P6="기타매출","R&amp;D용역",""))'),
    'U': (8, 'IF($C8="","",IF($O8="매출",$J8-$I8,$I8-$J8))'),
    'V': (8, 'IF($C8="","",MONTH($C8))'),
}


def _수식값미리계산(새, 이어받음, 매핑표, 날짜):
    """K~V 수식이 낼 값을 파이썬으로 미리 구해 칸에 같이 적어 둡니다.

    ※ 엑셀은 수식만 있고 계산값이 없으면, 다시 계산하기 전까지 그 칸을 빈칸으로 보여 줍니다.
      원장이 1만 줄이라 계산이 오래 걸려서, 내려받아 열면 K~T 가 한동안 비어 보였습니다.
      그래서 값을 미리 넣어 둡니다 (수식은 그대로라 나중에 다시 계산해도 같은 값입니다).
    """
    영업외 = {'이자비용', '잡손실', '잡이익', '외환차익', '외환차손',
              '외화환산이익', '외화환산손실', '유형자산처분손실'}
    계정 = _글자칸(새, 'Distribution account')
    차변 = pd.to_numeric(새.get('Debit'), errors='coerce').fillna(0).tolist()
    대변 = pd.to_numeric(새.get('Credit'), errors='coerce').fillna(0).tolist()
    값 = {c: [] for c in 자동수식열}
    for i in range(len(새)):
        d = 날짜[i]
        비었다 = d is None
        날 = None if 비었다 else 엑셀기준일 + datetime.timedelta(days=d)
        b, k = 계정[i], 계정[i].lower()
        bs, kor = 매핑표.get(b, ('', ''))
        분류 = ''
        if bs == 'IS':
            분류 = ('매출' if kor in ('제품매출', '기타매출')
                    else '매출원가' if kor == '제품매출원가'
                    else '영업외손익' if kor in 영업외 else '판관비')
        값['K'].append('' if 비었다 else 대변[i] - 차변[i])
        값['L'].append('' if 비었다 else f'{날.year}년')
        값['M'].append('' if 비었다 else f'{날.month}월')
        값['N'].append(bs if b else '')
        값['O'].append(분류)
        값['P'].append(kor if b else '')
        값['U'].append('' if 비었다 else
                       (대변[i] - 차변[i] if 분류 == '매출' else 차변[i] - 대변[i]))
        값['V'].append('' if 비었다 else 날.month)
    # Q·R 은 매출 계정이면 수식이 값을 만들어 냅니다 — 그 값도 같이 적어 둡니다
    Q, R = [], []
    for i in range(len(새)):
        kor = 매핑표.get(계정[i], ('', ''))[1]
        Q.append('친환경 바이오 소재' if kor == '제품매출'
                 else '기술 서비스 및 자문' if kor == '기타매출' else '')
        R.append('친환경 빨대/원자재 및 부자재 공급' if kor == '제품매출'
                 else 'R&D용역' if kor == '기타매출' else '')
    return 값, Q, R


def _값칸(열, r, sa, 수식, v):
    """수식과 그 계산값을 함께 적습니다 (글자 결과는 t="str")."""
    if isinstance(v, str):
        return (f'<c r="{열}{r}"{sa} t="str">{수식}<v>{_xml(v)}</v></c>'
                if v else f'<c r="{열}{r}"{sa} t="str">{수식}<v></v></c>')
    if isinstance(v, int) and not isinstance(v, bool):
        return f'<c r="{열}{r}"{sa}>{수식}<v>{v}</v></c>'
    return f'<c r="{열}{r}"{sa}>{수식}<v>{float(v)!r}</v></c>'


def _원장시트XML(원본xml, 새, 이어받음, 기간글, 매핑표=None):
    앞, 뒤, 머리행, 그룹서식, 자료서식, 수식본 = _시트본뜨기(원본xml)
    # 전월 파일에서 수식이 지워진 칸은 본디 수식으로 되살립니다
    되살림 = [c for c in 기본수식본 if c not in 수식본]
    수식본 = {**{c: 기본수식본[c] for c in 되살림}, **수식본}
    _되살린열.clear()
    _되살린열.extend(sorted(되살림))

    앞부분 = []
    for r in sorted(머리행):
        조각 = 머리행[r]
        if r == 3 and 기간글:                     # 맨 위 기간 표시만 새로 씁니다
            조각 = re.sub(
                r'<c r="A3"([^>]*?)(?:/>|>.*?</c>)',
                lambda m: '<c r="A3"' + re.sub(r'\s*t="\w+"', '', m.group(1))
                + f' t="inlineStr"><is><t>{_xml(기간글)}</t></is></c>',
                조각, count=1, flags=re.S)
        앞부분.append(조각)

    시작행, 끝행 = 6, 6 + len(새) - 1
    si = {c: i for i, c in enumerate(자동수식열)}
    칸값 = {col: (새[이름].tolist() if 이름 in 새.columns else [None] * len(새))
            for col, 이름 in 원장열자리.items()}
    날짜 = [_날짜일련(v) for v in 칸값['C']]
    미리값, 미리Q, 미리R = ({}, None, None)
    if 매핑표:
        미리값, 미리Q, 미리R = _수식값미리계산(새, 이어받음, 매핑표, 날짜)

    줄들, 첫줄 = [], {c: True for c in 자동수식열}
    for i in range(len(새)):
        r = 시작행 + i
        b = 칸값['B'][i]
        그룹행 = (b is None or (isinstance(b, float) and pd.isna(b))
                  or str(b).strip() in ('', 'nan'))
        서식 = 그룹서식 if 그룹행 else 자료서식
        칸 = []
        for col in 전체열:
            s = 서식.get(col) or 자료서식.get(col) or 그룹서식.get(col)
            sa = f' s="{s}"' if s else ''
            if col in 원장열자리:
                v = 칸값[col][i]
                if col == 'C':
                    d = 날짜[i]
                    칸.append(f'<c r="{col}{r}"{sa}><v>{d}</v></c>' if d is not None
                              else f'<c r="{col}{r}"{sa}/>')
                elif v is None or (isinstance(v, float) and pd.isna(v)):
                    칸.append(f'<c r="{col}{r}"{sa}/>')
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    칸.append(f'<c r="{col}{r}"{sa}><v>{v!r}</v></c>')
                else:
                    칸.append(f'<c r="{col}{r}"{sa} t="inlineStr">'
                              f'<is><t xml:space="preserve">{_xml(v)}</t></is></c>')
            elif col in 자동수식열 and col in 수식본:
                원행, 본 = 수식본[col]
                if 첫줄[col]:
                    첫줄[col] = False
                    식 = (f'<f t="shared" ref="{col}{시작행}:{col}{끝행}" '
                          f'si="{si[col]}">{_수식옮기기(본, 원행, r)}</f>')
                else:
                    식 = f'<f t="shared" si="{si[col]}"/>'
                if col in 미리값:
                    칸.append(_값칸(col, r, sa, 식, 미리값[col][i]))
                else:
                    칸.append(f'<c r="{col}{r}"{sa}>{식}</c>')
            elif col in 이어받기열:
                v = 이어받음.get(col, [None] * len(새))[i]
                if v:
                    칸.append(f'<c r="{col}{r}"{sa} t="inlineStr">'
                              f'<is><t xml:space="preserve">{_xml(v)}</t></is></c>')
                elif col in 수식본:
                    원행, 본 = 수식본[col]
                    식 = f'<f>{_수식옮기기(본, 원행, r)}</f>'
                    미리 = ({'Q': 미리Q, 'R': 미리R}.get(col) or [None] * len(새))[i]
                    칸.append(_값칸(col, r, sa, 식, 미리)
                              if 미리 is not None
                              else f'<c r="{col}{r}"{sa}>{식}</c>')
                else:
                    칸.append(f'<c r="{col}{r}"{sa}/>')
            else:
                칸.append(f'<c r="{col}{r}"{sa}/>')
        줄들.append(f'<row r="{r}" spans="1:22" ht="15" customHeight="1">'
                    + ''.join(칸) + '</row>')

    앞 = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:V{끝행}"/>', 앞, count=1)
    뒤 = re.sub(r'(<autoFilter ref="A5:)[A-Z]+\d+(")', rf'\g<1>T{끝행}\g<2>', 뒤, count=1)
    return 앞 + '<sheetData>' + ''.join(앞부분) + ''.join(줄들) + '</sheetData>' + 뒤



def _관계사전(rel):
    """workbook.xml.rels → {rId: 'xl/…' 경로}. 속성 순서(Id·Type·Target)와 절대/상대 Target 모두 받습니다.
    (엑셀은 Id 가 먼저, openpyxl 로 저장한 파일은 Type·Target·Id 순서에 Target 이 '/xl/…' 절대 경로)"""
    나온것 = {}
    for m in re.finditer(r'<Relationship\b[^>]*?/>', rel):
        조각 = m.group(0)
        아이디 = re.search(r'\bId="([^"]+)"', 조각)
        대상 = re.search(r'\bTarget="([^"]+)"', 조각)
        if not (아이디 and 대상):
            continue
        t = 대상.group(1)
        나온것[아이디.group(1)] = t.lstrip('/') if t.startswith('/') else posixpath.normpath('xl/' + t)
    return 나온것

def _시트지도(z):
    """시트 이름 → XML 파일 경로 (엑셀 안에서의 순서 그대로)."""
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rel = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid = _관계사전(rel)
    시트 = []
    for m in re.finditer(r'<sheet\b([^>]*?)/>', wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        시트.append((_풀기(str(a.get('name'))), rid.get(a.get('r:id'), 'xl/')))
    return 시트


def _전월원장시트(z):
    """전월 파일에서 「○○년 원장_raw」 시트의 이름과 XML 파일 경로를 찾습니다."""
    시트 = _시트지도(z)
    원장 = sorted((n, f) for n, f in 시트 if '원장_raw' in str(n))
    if not 원장:
        raise ValueError('전월 실적보고자료에서 「○○년 원장_raw」 시트를 찾지 못했습니다. '
                         '지난달에 쓰시던 실적보고 엑셀이 맞는지 확인해 주세요.')
    return 원장[-1], [n for n, _f in 시트]


def _관계표(z, 파트):
    """파트 하나가 딸린 파트들을 무엇무엇 쓰고 있는지 → (관계파일 이름, {rId: 파트경로})."""
    관계파일 = posixpath.join(posixpath.dirname(파트), '_rels',
                              posixpath.basename(파트) + '.rels')
    try:
        글 = z.read(관계파일).decode('utf-8')
    except KeyError:
        return 관계파일, {}
    나온것 = {}
    for m in re.finditer(r'<Relationship\b[^>]*?/>', 글):
        조각 = m.group(0)
        if 'TargetMode="External"' in 조각:
            continue
        아이디 = re.search(r'Id="([^"]+)"', 조각)
        대상 = re.search(r'Target="([^"]+)"', 조각)
        if not (아이디 and 대상):
            continue
        나온것[아이디.group(1)] = posixpath.normpath(
            posixpath.join(posixpath.dirname(파트), 대상.group(1).lstrip('/')))
    return 관계파일, 나온것


def _딸린파트(z, 시작):
    """그 파트와, 그 파트만 타고 내려가 만나는 모든 파트 (피벗 → 캐시 → 기록, 그림 → 그림파일)."""
    본것, 남은것 = set(), [시작]
    while 남은것:
        p = 남은것.pop()
        if p in 본것:
            continue
        본것.add(p)
        남은것.extend(_관계표(z, p)[1].values())
    return 본것


def _시트빼기(z, 뺄이름):
    """실적보고 엑셀에서 시트 하나를 통째로 들어냅니다 → (지울 파트 set, {파트: 새 바이트}).

    ※ 「당월 실적집계」처럼 이제 쓰지 않는 시트를 내려받는 파일에서 빼기 위한 것입니다.
      그 시트에만 딸린 피벗표·피벗캐시·그림·메모까지 같이 빼되, 다른 시트도 같이 쓰는
      파트는 그대로 둡니다. 워크북 목차·관계·파일목록·문서정보도 함께 맞춰 고칩니다.
      시트가 없으면 (set(), {}) 를 돌려주므로 그냥 지나갑니다.
    """
    wb = z.read('xl/workbook.xml').decode('utf-8')
    wbrel = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    묶음 = _관계사전(wbrel)

    시트조각, 자리, 뺄rId, 뺄파트 = None, -1, None, None
    for i, m in enumerate(re.finditer(r'<sheet\b[^>]*?/>', wb)):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(0)))
        if _풀기(str(a.get('name'))) == 뺄이름:
            시트조각, 자리, 뺄rId = m.group(0), i, a.get('r:id')
            뺄파트 = posixpath.normpath(묶음.get(뺄rId, ''))
            break
    if 시트조각 is None or not 뺄파트:
        return set(), {}

    모든파트 = set(z.namelist())
    후보 = _딸린파트(z, 뺄파트) & 모든파트

    # 살아남는 시트가 같이 쓰는 파트는 되살립니다 (워크북 자신은 아래에서 따로 고칩니다)
    바뀜 = True
    while 바뀜:
        바뀜 = False
        for 파트 in sorted(모든파트 - 후보):
            if 파트 == 'xl/workbook.xml' or 파트.endswith('.rels'):
                continue
            for t in _관계표(z, 파트)[1].values():
                if t in 후보:
                    후보 -= _딸린파트(z, t)
                    바뀜 = True
                    break
            if 바뀜:
                break

    지울파트 = set()
    for 파트 in 후보:
        지울파트.add(파트)
        관계파일 = _관계표(z, 파트)[0]
        if 관계파일 in 모든파트:
            지울파트.add(관계파일)

    # ── 워크북 목차 고치기 ────────────────────────────────
    새wb = wb.replace(시트조각, '', 1)

    뺄rId들 = {뺄rId}
    for m in re.finditer(r'<pivotCache\b[^>]*?/>', 새wb):
        아이디 = re.search(r'r:id="([^"]+)"', m.group(0))
        if 아이디 and posixpath.normpath(묶음.get(아이디.group(1), '')) in 지울파트:
            새wb = 새wb.replace(m.group(0), '', 1)
            뺄rId들.add(아이디.group(1))
    새wb = re.sub(r'<pivotCaches>\s*</pivotCaches>', '', 새wb)

    # 시트 번호로 매인 이름(정의된 이름)은 앞시트가 빠진 만큼 번호를 당겨 줍니다
    def _이름고침(m):
        조각 = m.group(0)
        번호 = int(re.search(r'localSheetId="(\d+)"', 조각).group(1))
        if 번호 == 자리:
            return ''
        return 조각.replace(f'localSheetId="{번호}"', f'localSheetId="{번호 - 1}"') \
            if 번호 > 자리 else 조각
    새wb = re.sub(r'<definedName\b[^>]*localSheetId="\d+"[^>]*>.*?</definedName>',
                  _이름고침, 새wb, flags=re.S)
    새wb = re.sub(r'<definedNames>\s*</definedNames>', '', 새wb)

    def _보기고침(m):
        조각 = m.group(0)
        for 이름 in ('activeTab', 'firstSheet'):
            n = re.search(rf'{이름}="(\d+)"', 조각)
            if n:
                v = max(int(n.group(1)) - (1 if int(n.group(1)) > 자리 else 0), 0)
                조각 = 조각.replace(n.group(0), f'{이름}="{v}"')
        return 조각
    새wb = re.sub(r'<workbookView\b[^>]*?/>', _보기고침, 새wb, count=1)

    새rel = wbrel
    for rid in 뺄rId들:
        새rel = re.sub(rf'<Relationship[^>]*Id="{re.escape(rid)}"[^>]*/>', '', 새rel, count=1)

    고칠것 = {'xl/workbook.xml': 새wb.encode('utf-8'),
              'xl/_rels/workbook.xml.rels': 새rel.encode('utf-8')}

    # ── 파일목록([Content_Types]) 에서 빠진 파트를 지웁니다 ──
    ct = z.read('[Content_Types].xml').decode('utf-8')
    for 파트 in 지울파트:
        ct = re.sub(rf'<Override PartName="/{re.escape(파트)}"[^>]*/>', '', ct)
    고칠것['[Content_Types].xml'] = ct.encode('utf-8')

    # ── 문서정보(시트 이름 목록)도 한 칸 줄입니다 ─────────
    try:
        app = z.read('docProps/app.xml').decode('utf-8')
        표기 = 시트조각 and re.search(r'name="([^"]*)"', 시트조각).group(1)
        낱말 = f'<vt:lpstr>{표기}</vt:lpstr>'
        if 낱말 in app:
            app = app.replace(낱말, '', 1)
            app = re.sub(r'(<TitlesOfParts><vt:vector size=")(\d+)(")',
                         lambda m: m.group(1) + str(max(int(m.group(2)) - 1, 0)) + m.group(3),
                         app, count=1)
            app = re.sub(r'(<vt:lpstr>워크시트</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)(</vt:i4>)',
                         lambda m: m.group(1) + str(max(int(m.group(2)) - 1, 0)) + m.group(3),
                         app, count=1)
            고칠것['docProps/app.xml'] = app.encode('utf-8')
    except KeyError:
        pass

    return 지울파트, 고칠것


def _칸값바꾸기(시트xml, 칸, 값):
    """시트 XML 에서 칸 하나의 값만 갈아끼웁니다 (서식 s= 는 그대로 둡니다).

    ※ 「당월 실적집계」의 H2 가 몇 월 보고서인지 정하는 칸입니다.
      이 칸이 지난달 그대로면 표 전체가 지난달 숫자로 남습니다.
    """
    행 = re.search(r'\d+', 칸).group(0)
    새칸 = None

    def _고치기(m):
        속성 = re.sub(r'\s*t="\w+"', '', m.group(1))
        return f'<c r="{칸}"{속성}><v>{값}</v></c>'

    시트xml, 바뀜 = re.subn(rf'<c r="{칸}"([^>]*?)(?:/>|>.*?</c>)', _고치기,
                            시트xml, count=1, flags=re.S)
    if 바뀜:
        return 시트xml, True

    # 그 칸이 아예 없으면 해당 행의 「열 순서에 맞는 자리」에 새로 끼워 넣습니다
    열 = re.match(r'[A-Z]+', 칸).group(0)
    새칸 = f'<c r="{칸}"><v>{값}</v></c>'

    def _열번호(c):
        n = 0
        for ch in c:
            n = n * 26 + (ord(ch) - 64)
        return n

    행틀 = re.search(rf'<row[^>]*\br="{행}"[^>]*>(.*?)</row>', 시트xml, re.S)
    if not 행틀:
        return 시트xml, False
    속 = 행틀.group(1)
    자리 = len(속)
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"', 속):
        if _열번호(cm.group(1)) > _열번호(열):
            자리 = cm.start()
            break
    새속 = 속[:자리] + 새칸 + 속[자리:]
    return (시트xml[:행틀.start(1)] + 새속 + 시트xml[행틀.end(1):]), True


# ══════════════════════════════════════════════════════════════
#  「26년 자금」 시트 — 왼쪽은 원장에서 뽑은 실제 현금흐름, 오른쪽은 남은 달의 예상 출금
#  ※ 보고 월이 바뀌면 왼쪽은 그달까지 다시 쓰고, 오른쪽은 그다음 달부터만 남깁니다.
# ══════════════════════════════════════════════════════════════
자금항목이름 = {'인건비성 항목': '인건비성 항목', '커미션 비용': '커미션 비용',
                '외부용역': '외부용역', '건물': '건물관리비', '차량': '차량유지비',
                'ICT': 'ICT', '출장비': '출장비', '접대비': '접대비',
                '소모성경비': '소모성경비', '기타관리경비': '기타관리경비',
                '기타': '기타관리경비', '샘플 및 금형관련비': '샘플 및 금형관련비',
                '금형관련비': '샘플 및 금형관련비', '운반 및 창고료': '운반 및 창고료',
                '정부지원': '정부지원'}
자금분류표 = {'AR': {'매출채권', '미수금', '선수금'},
              'AP': {'매입채무', '미지급금', '미지급비용'},
              'PAY': {'예수금', '예수금(단체보험)', '급여', '복리후생비', '퇴직급여'},
              'EQ': {'보통주자본금', '주식발행초과금', '자본잉여금', '기타자본'},
              'DEBT': {'단기차입금', '장기차입금', '전환사채', '단기대여금', '장기대여금'},
              'LEASE': {'리스부채(유동)', '리스부채(비유동)'},
              'CAPEX': {'기계장치', '건설중인자산', '시설장치', '비품', '공구와기구',
                        '무형자산', '특허권', '사용권자산', '사용권자산(비유동)'},
              'FX': {'외환차손익', '외화환산손익'}}
자금구분 = {'AR': '영업활동', 'ETCIN': '영업활동', 'AP': '영업활동', 'PAY': '영업활동',
            'OTHOP': '영업활동', 'CAPEX': '투자활동', 'EQ': '재무활동', 'DEBT': '재무활동',
            'LEASE': '재무활동', 'FX': '환율변동 효과'}
자금이름 = {'AR': '매출채권 회수', 'ETCIN': '기타입금', 'AP': '매입채무 지급',
            'PAY': '급여', 'OTHOP': '기타 비용',
            'CAPEX': '설비 투자', 'EQ': '유상증자', 'DEBT': '차입증감',
            'LEASE': '리스료 지급', 'FX': '환율변동 효과'}
자금표시 = 252          # 「이 파일에는 자금 시트 서식이 이미 들어 있다」는 표시로 쓰는 번호


def _열이름(n):
    이름 = ''
    while n:
        n, 나머지 = divmod(n - 1, 26)
        이름 = chr(65 + 나머지) + 이름
    return 이름


def _열번호(이름):
    n = 0
    for 글 in str(이름):
        n = n * 26 + (ord(글) - 64)
    return n


def _자금칸그대로(z, 파일):
    """옛 자금 시트의 칸을 XML 그대로 {(행, 열번호): 조각} 으로 읽습니다.

    ※ 담당자가 손으로 넣은 3번 표(예상 입금)를 한 글자도 건드리지 않고
      그대로 옮기기 위한 것입니다.
    """
    try:
        x = z.read(파일).decode('utf-8')
    except (KeyError, AttributeError):
        return {}
    if '<sheetData>' not in x:
        return {}
    본문 = x.split('<sheetData>', 1)[1].rsplit('</sheetData>', 1)[0]
    표 = {}
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"[^>]*?(?:/>|>.*?</c>)', 본문, re.S):
        표[(int(m.group(2)), _열번호(m.group(1)))] = m.group(0)
    return 표


def _자금코드(계정과목):
    for k, s in 자금분류표.items():
        if 계정과목 in s:
            return k
    return 'OTHOP'


def _자금서식(스타일):
    """자금 시트에 필요한 서식을 styles.xml 뒤에 덧붙입니다.

    ※ 반드시 목록의 「끝」에 붙여야 합니다. 앞에 끼워 넣으면 기존 번호가 전부
      한 칸씩 밀려 파일 전체 서식이 망가집니다.
    ※ 이미 붙어 있으면(자금표시 번호를 쓰는 xf 가 있으면) 그 번호를 그대로 씁니다.
    """
    xf목록 = re.search(r'<cellXfs count="\d+">(.*?)</cellXfs>', 스타일, re.S)
    if xf목록:
        칸들 = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', xf목록.group(1), re.S)
        for i, x in enumerate(칸들):
            if f'numFmtId="{자금표시}"' in x:
                return 스타일, [i + k for k in range(17)], False

    글꼴수 = int(re.search(r'<fonts count="(\d+)"', 스타일).group(1))
    테두리수 = int(re.search(r'<borders count="(\d+)"', 스타일).group(1))
    xf수 = int(re.search(r'<cellXfs count="(\d+)"', 스타일).group(1))
    쓰인번호 = {int(v) for v in re.findall(r'<numFmt numFmtId="(\d+)"', 스타일)}
    F날짜 = next(n for n in range(자금표시 + 1, 400) if n not in 쓰인번호)
    F숫자 = next(n for n in range(F날짜 + 1, 400) if n not in 쓰인번호)

    새서식 = (f'<numFmt numFmtId="{F날짜}" formatCode="yyyy\\-mm\\-dd"/>'
              f'<numFmt numFmtId="{F숫자}" formatCode="#,##0.00;[Red]\\({{#,##0.00}}\\)"/>'
              f'<numFmt numFmtId="{자금표시}" formatCode="General"/>')
    if '<numFmts' in 스타일:
        서식수 = int(re.search(r'<numFmts count="(\d+)"', 스타일).group(1))
        스타일 = 스타일.replace('</numFmts>', 새서식 + '</numFmts>', 1)
        스타일 = 스타일.replace(f'<numFmts count="{서식수}"',
                                f'<numFmts count="{서식수 + 3}"', 1)
    else:
        스타일 = 스타일.replace('<fonts', f'<numFmts count="3">{새서식}</numFmts><fonts', 1)

    맑은 = '<name val="맑은 고딕"/><family val="2"/>'
    새글꼴 = (f'<font><sz val="12"/><b/><color rgb="FF000000"/>{맑은}</font>'
              f'<font><sz val="11"/><b/><color rgb="FF000000"/>{맑은}</font>'
              f'<font><sz val="9"/><color rgb="FF3B6AA0"/>{맑은}</font>'
              f'<font><sz val="10"/><color rgb="FF000000"/>{맑은}</font>'
              f'<font><sz val="10"/><b/><color rgb="FF000000"/>{맑은}</font>')
    스타일 = 스타일.replace('</fonts>', 새글꼴 + '</fonts>', 1)
    스타일 = 스타일.replace(f'<fonts count="{글꼴수}"', f'<fonts count="{글꼴수 + 5}"', 1)
    G제목, G머리, G설명, G본문, G굵게 = (글꼴수, 글꼴수 + 1, 글꼴수 + 2,
                                        글꼴수 + 3, 글꼴수 + 4)

    점 = '<left style="dotted"><color rgb="FF808080"/></left>'
    실 = '<bottom style="thin"><color rgb="FF000000"/></bottom>'
    위실 = '<top style="thin"><color rgb="FF000000"/></top>'
    새테두리 = ('<border><left/><right/><top/><bottom/><diagonal/></border>'
                f'<border>{점}<right/><top/><bottom/><diagonal/></border>'
                f'<border><left/><right/>{위실}{실}<diagonal/></border>'
                f'<border>{점}<right/>{위실}{실}<diagonal/></border>'
                f'<border><left/><right/><top/>{실}<diagonal/></border>'
                f'<border>{점}<right/><top/>{실}<diagonal/></border>')
    스타일 = 스타일.replace('</borders>', 새테두리 + '</borders>', 1)
    스타일 = 스타일.replace(f'<borders count="{테두리수}"',
                            f'<borders count="{테두리수 + 6}"', 1)
    B없음, B점, B머리0, B머리, B끝0, B끝 = (테두리수, 테두리수 + 1, 테두리수 + 2,
                                           테두리수 + 3, 테두리수 + 4, 테두리수 + 5)

    def xf(글꼴, 테두리, 서식=0, 가로=None):
        맞춤 = f'<alignment horizontal="{가로}" vertical="center"/>' if 가로 else ''
        속성 = (f'<xf numFmtId="{서식}" fontId="{글꼴}" fillId="0" borderId="{테두리}" '
                f'xfId="0" applyFont="1" applyBorder="1"'
                + (' applyNumberFormat="1"' if 서식 else '')
                + (' applyAlignment="1"' if 가로 else ''))
        return 속성 + (f'>{맞춤}</xf>' if 맞춤 else '/>')

    새xf = [
        xf(G제목, B없음, 자금표시),                    # 0 제목 (← 여기가 표시 번호)
        xf(G굵게, B없음), xf(G설명, B없음),            # 1 소제목 · 2 설명
        xf(G머리, B머리0, 가로='center'),              # 3 머리(첫 칸)
        xf(G머리, B머리, 가로='center'),               # 4 머리
        xf(G본문, B없음, 서식=F날짜, 가로='center'),   # 5 일자(첫 칸)
        xf(G본문, B점, 서식=F숫자),                    # 6 숫자
        xf(G본문, B점, 가로='center'),                 # 7 가운데
        xf(G본문, B점, 가로='left'),                   # 8 왼쪽
        xf(G본문, B없음, 서식=F날짜, 가로='center'),   # 9 (예비)
        xf(G본문, B끝0), xf(G본문, B끝),               # 10 끝줄(첫 칸) · 11 끝줄
        xf(G설명, B없음, 가로='left'),                 # 12 주석
        xf(G본문, B점, 서식=F날짜, 가로='center'),     # 13 일자
        xf(G본문, B없음, 가로='left'),                 # 14 근거
        xf(G굵게, B없음, 가로='left'),                 # 15 근거 제목
        xf(G본문, B없음, 서식=F숫자),                  # 16 근거 숫자
    ]
    스타일 = 스타일.replace('</cellXfs>', ''.join(새xf) + '</cellXfs>', 1)
    스타일 = 스타일.replace(f'<cellXfs count="{xf수}"',
                            f'<cellXfs count="{xf수 + len(새xf)}"', 1)
    return 스타일, [xf수 + i for i in range(len(새xf))], True


def _예상기준행찾기(z, 파일, 공유):
    """「월별 실적집계」에서 예상 자금을 뽑아 쓸 기준 줄의 행 번호를 찾습니다.

    돌려주는 값 : {'매출': r, '원가': r, '관리비': r, '상각': [r…],
                   '인건비': [r…], 'SRED': [r…]}
      · 매출 · 원가 · 관리비 : C열의 「매출액 합계」 · 「매출원가 합계」 · 「관리비 합계」
      · 상각 · 인건비        : C열의 「n) …」 항목 중 이름에 「상각비」 · 「인건비」가 든 줄
      · SRED                 : E열 세부 줄 중 이름에 「SR&ED」 가 든 줄 (일회성 보조금)
    항목 이름만 보고 찾으므로 줄이 위아래로 밀려도 그대로 따라갑니다.
    """
    표 = _시트칸읽기(z, 파일, ('C', 'E'), 공유)
    씨 = {행: str(칸.get('C', '')).strip() for 행, 칸 in 표.items()}
    이 = {행: str(칸.get('E', '')).strip() for 행, 칸 in 표.items()}
    찾음 = {'상각': [], '인건비': [], 'SRED': []}
    for 이름, 키 in (('매출액 합계', '매출'), ('매출원가 합계', '원가'),
                     ('관리비 합계', '관리비')):
        찾음[키] = next((r for r in sorted(씨) if 씨[r].startswith(이름)), None)
    관리비행 = 찾음.get('관리비')
    끝 = next((r for r in sorted(씨) if 관리비행 and r > 관리비행
               and 씨[r].startswith('영업이익')), None)
    for r in sorted(씨):
        if 관리비행 is None or r <= 관리비행 or (끝 and r >= 끝):
            continue
        m = re.match(r'^(\d+)\)\s*(.+)$', 씨[r])
        if m:
            이름 = m.group(2)
            if '상각비' in 이름:
                찾음['상각'].append(r)
            elif '인건비' in 이름:
                찾음['인건비'].append(r)
    for r in sorted(이):
        if 관리비행 is not None and 관리비행 < r < (끝 or r + 1) \
                and re.search(r'SR\s*&?\s*ED', 이[r], re.I):
            찾음['SRED'].append(r)
    return 찾음


def _관리비항목찾기(z, 파일, 공유):
    """「월별 실적집계」 C열에서 관리비 하위 항목(1) … 13))과 그 행 번호를 찾습니다."""
    표 = _시트칸읽기(z, 파일, ('C',), 공유)
    글 = {행: str(칸.get('C', '')).strip() for 행, 칸 in 표.items()}
    시작 = next((r for r in sorted(글) if 글[r].startswith('관리비 합계')), None)
    끝 = next((r for r in sorted(글) if r > (시작 or 0)
               and 글[r].startswith('영업이익')), None)
    if 시작 is None:
        return []
    항목 = []
    for 행 in sorted(글):
        if 행 <= 시작 or (끝 and 행 >= 끝):    # 매출·매출원가와 아래쪽 요약표는 건너뜁니다
            continue
        m = re.match(r'^(\d+)\)\s*(.+)$', 글[행])
        if not m:
            continue
        이름 = m.group(2).strip()
        if '상각비' in 이름:                   # 현금이 나가지 않으니 예상 출금에서 뺍니다
            continue
        항목.append((이름, 행))
    return 항목


def _자금시트만들기(z, 시트칸, 공유, 새, 이어받음, 매핑표, 연도, 보고달):
    """「26년 자금」 시트 XML 을 통째로 다시 만듭니다. → (시트xml, 새styles, 안내)"""
    자금시트 = next((n for n in 시트칸 if '자금' in n), None)
    if not 자금시트:
        return None, None, None, {}
    현금계정 = {k.strip().lower() for k, (b, kor) in 매핑표.items()
                if b == 'BS' and ('현금' in kor or '예금' in kor)}
    한글 = {k.strip().lower(): v[1] for k, v in 매핑표.items()}
    비에스 = {k.strip().lower(): v[0] for k, v in 매핑표.items()}

    g = pd.DataFrame({
        '그룹': pd.Series(_글자칸(새, 원장열자리['A'])).replace('', None).ffill().fillna(''),
        '계정': _글자칸(새, 'Distribution account'),
        '전표': _글자칸(새, 'Transaction ID'),
        'Description': _글자칸(새, 'Description'),
        '#': _글자칸(새, '#'),
        'Name': _글자칸(새, 'Name'),
        '일련': [_날짜일련(v) for v in 새['Transaction date']],
        '차변': pd.to_numeric(새.get('Debit'), errors='coerce').fillna(0).tolist(),
        '대변': pd.to_numeric(새.get('Credit'), errors='coerce').fillna(0).tolist(),
        '잔액': pd.to_numeric(새.get('Balance'), errors='coerce').fillna(0).tolist(),
        '활동': [v or '' for v in 이어받음.get('Q', [None] * len(새))],
    })
    # 기초 현금 = 현금·예금 계정의 Beginning Balance 합계
    기초현금액 = float(g.loc[g['계정'].eq('Beginning Balance')
                             & g['그룹'].str.lower().isin(현금계정), '잔액'].sum())

    g = g[g['계정'].ne('') & g['일련'].notna()].copy()
    키 = g['계정'].str.lower()
    g['계정과목'] = 키.map(한글).fillna('')
    g['계정분류'] = 키.map(비에스).fillna('')
    g['현금'] = 키.isin(현금계정)

    상대 = g[~g['현금']].copy()
    상대['코드'] = [_자금코드(n) for n in 상대['계정과목']]
    상대.loc[상대['계정과목'].str.contains('감가상각누계', na=False), '코드'] = 'CAPEX'
    # SR&ED 환급금은 매출채권 회수가 아니라 기타입금으로 갈라 냅니다
    지원전표 = set(g.loc[_기타입금줄(g, g['현금'], g['차변']), '전표'])
    상대.loc[상대['전표'].isin(지원전표), '코드'] = 'ETCIN'
    상대['크기'] = (상대['대변'] - 상대['차변']).abs()
    상대['환율아님'] = (~상대['코드'].eq('FX')).astype(int)
    상대['세부'] = np.where(상대['계정분류'].eq('IS') & 상대['활동'].ne(''),
                            상대['활동'].map(자금항목이름).fillna('기타관리경비'), '')
    대표 = (상대.sort_values(['환율아님', '크기'], ascending=False)
            .drop_duplicates('전표').set_index('전표')[['코드', '세부']])

    현 = g[g['현금']].sort_values(['일련', '전표'], kind='stable').reset_index(drop=True)
    현['코드'] = 현['전표'].map(대표['코드']).fillna('이체')
    현['세부'] = 현['전표'].map(대표['세부']).fillna('')
    현['구분'] = 현['코드'].map(자금구분).fillna('현금 간 이체')
    현['항목'] = np.where(현['세부'].ne(''), 현['세부'],
                          현['코드'].map(자금이름).fillna('현금 계정 간 이체'))
    기말액 = 기초현금액 + float((현['차변'] - 현['대변']).sum())

    # ── 담당자가 손으로 채운 3번 표(예상 입금)는 그대로 옮깁니다 ──
    옛칸 = _자금칸그대로(z, 시트칸[자금시트])
    옛글 = _시트칸읽기(z, 시트칸[자금시트],
                       tuple(_열이름(i) for i in range(1, 31)), 공유)
    옛제목 = {}
    보존시작 = None
    for 행, 칸 in 옛글.items():
        if 행 > 8:
            continue
        for 열, v in 칸.items():
            m = re.match(r'^\s*([23])\s*\.', str(v))
            if m:
                옛제목.setdefault(m.group(1), (행, _열번호(열), str(v)))
                if m.group(1) == '3':
                    보존시작 = _열번호(열) - 1      # 3번 표 앞 빈 칸부터 오른쪽 전부
    보존칸 = ({칸: 조각 for 칸, 조각 in 옛칸.items() if 칸[1] >= 보존시작}
              if 보존시작 else {})

    스타일, S, 새서식 = _자금서식(z.read('xl/styles.xml').decode('utf-8'))
    (S제목, S소제목, S설명, S머리0, S머리, S일자0, S숫자, S가운데, S왼쪽,
     _예비, S끝0, S끝, _주석, S일자, S근거, S근거제목, S근거숫자) = S

    칸들 = []

    def 글(ref, s, v):
        칸들.append(f'<c r="{ref}" s="{s}" t="inlineStr">'
                    f'<is><t xml:space="preserve">{_xml(v)}</t></is></c>')

    def 숫(ref, s, v):
        # ※ numpy 숫자를 그대로 쓰면 'np.float64(…)' 로 적혀 파일이 깨집니다
        칸들.append(f'<c r="{ref}" s="{s}"><v>{float(v)!r}</v></c>')

    def 식(ref, s, f):
        칸들.append(f'<c r="{ref}" s="{s}"><f>{_xml(f)}</f></c>')

    def 빈(ref, s):
        칸들.append(f'<c r="{ref}" s="{s}"/>')

    머리 = ['일자', '차변', '대변', '기말', '구분', '활동성항목']   # 1. 실제 (기말 있음)
    예상머리 = ['일자', '차변', '대변', '구분', '활동성항목']        # 2·3. 예상 (기말 없음)
    왼열 = ['B', 'C', 'D', 'E', 'F', 'G']
    오열 = ['I', 'J', 'K', 'L', 'M']               # 2. 예상 출금 (대변만 클로드가 채웁니다)
    입열 = ['O', 'P', 'Q', 'R', 'S']               # 3. 예상 입금 (담당자가 직접 넣습니다)
    글('A1', S제목, '● 원장 작성 현금성 자산 계정별원장')
    글('B3', S소제목, '1. 실제 현금성 자금 흐름')
    글('B4', S설명, f'- 원장상 현금흐름을 기준으로 클로드가 작성한 현금흐름 '
                    f'({연도}년 1~{보고달}월)')
    글('I3', S소제목, 옛제목.get('2', (0, 0, '2. 예상 현금성 자금 흐름(출금)'))[2])
    글('I4', S설명, f'- ※ 가정입니다(실적 아님) — 매출액 · 매출원가 · 판관비(현금) 세 가지를 '
                    f'1~{보고달}월 실적 월평균으로 놓고 클로드가 채운 값 ({보고달 + 1}~12월)'
                    if 보고달 < 12 else '- 12월까지 실적이 다 들어와 예상분은 없습니다')
    for 열, 이름 in zip(왼열, 머리):
        글(f'{열}5', S머리0 if 열 == 'B' else S머리, 이름)
    for 열, 이름 in zip(오열, 예상머리):
        글(f'{열}5', S머리0 if 열 == 'I' else S머리, 이름)

    # ── 왼쪽 : 원장에서 뽑은 실제 현금 계정별원장
    r = 6
    빈(f'B{r}', S일자0)
    빈(f'C{r}', S숫자)
    빈(f'D{r}', S숫자)
    숫(f'E{r}', S숫자, round(기초현금액, 2))
    글(f'F{r}', S가운데, '기초')
    글(f'G{r}', S왼쪽, f'{연도}년 1월 1일 현금·예금 잔액')
    for _i, t in 현.iterrows():
        r += 1
        숫(f'B{r}', S일자, int(t['일련']))
        차, 대 = float(t['차변']), float(t['대변'])
        숫(f'C{r}', S숫자, round(차, 2)) if 차 else 빈(f'C{r}', S숫자)
        숫(f'D{r}', S숫자, round(대, 2)) if 대 else 빈(f'D{r}', S숫자)
        식(f'E{r}', S숫자, f'E{r - 1}+C{r}-D{r}')
        글(f'F{r}', S가운데, str(t['구분']))
        글(f'G{r}', S왼쪽, str(t['항목']))
    왼끝 = r + 1
    for 열 in 왼열:
        빈(f'{열}{왼끝}', S끝0 if 열 == 'B' else S끝)

    # ── 오른쪽 : 남은 달의 예상 현금 증감 (보고월 다음 달부터 12월까지)
    #    가정은 셋뿐입니다 — 매출액 · 매출원가 · 판관비(현금). 셋 다 1~보고달 실적의
    #    월평균이라, 실적이 바뀌면 =AVERAGE(…) 수식을 타고 예상액도 같이 바뀝니다.
    #      · 매출채권 회수 = 그 달 매출액            (운전자본은 지금 수준 유지라고 봅니다)
    #      · 매입채무 지급 = 그 달 매출원가
    #      · 급여          = 인건비 + 기타 인건비성
    #      · 기타 비용     = 관리비 합계 − 상각비 − SR&ED − 급여
    #    상각비는 현금이 나가지 않아 빼고, SR&ED 보조금은 한 번 받고 마는 돈이라 뺍니다.
    집계시트 = next((n for n in 시트칸 if '월별' in n and '실적집계' in n), None)
    기준행 = _예상기준행찾기(z, 시트칸[집계시트], 공유) if 집계시트 else {}
    예상월 = list(range(보고달 + 1, 13))
    끝열 = _열이름(23 + 보고달)          # X=1월 … 1~보고달 평균을 냅니다
    r = 6
    빈(f'I{r}', S일자0)
    빈(f'J{r}', S숫자)
    빈(f'K{r}', S숫자)
    글(f'L{r}', S가운데, '기초')
    글(f'M{r}', S왼쪽, f'{연도}년 {보고달}월말 실제 잔액')

    # 수식이 낼 값을 미리 구해 같이 적어 둡니다 (엑셀을 열자마자 보이고, 대시보드도 읽습니다)
    달칸 = [_열이름(24 + k) for k in range(보고달)]          # X=1월 …
    읽은 = (_시트칸읽기(z, 시트칸[집계시트], tuple(달칸), 공유)
            if 집계시트 else {})

    def 줄평균(행들):
        """그 줄들의 1~보고달 합을 보고달로 나눈 월평균."""
        모두 = 0.0
        for 행 in ([행들] if isinstance(행들, int) else (행들 or [])):
            for c in 달칸:
                모두 += float(pd.to_numeric(pd.Series([읽은.get(행, {}).get(c)]),
                                            errors='coerce').fillna(0).iat[0])
        return 모두 / max(보고달, 1)

    def 평균식(행들):
        """=AVERAGE(…)+AVERAGE(…) — 줄이 여럿이면 더합니다. 없으면 0."""
        목록 = [행들] if isinstance(행들, int) else list(행들 or [])
        if not 목록 or not 집계시트:
            return '0'
        return '+'.join(f"AVERAGE('{집계시트}'!$X${h}:${끝열}${h})" for h in 목록)

    매출식, 원가식 = 평균식(기준행.get('매출')), 평균식(기준행.get('원가'))
    급여식 = 평균식(기준행.get('인건비'))
    기타식 = (f"{평균식(기준행.get('관리비'))}-({평균식(기준행.get('상각'))})"
              f"-({평균식(기준행.get('SRED'))})-({급여식})")
    월매출 = 줄평균(기준행.get('매출'))
    월원가 = 줄평균(기준행.get('원가'))
    월급여 = 줄평균(기준행.get('인건비'))
    월기타 = (줄평균(기준행.get('관리비')) - 줄평균(기준행.get('상각'))
              - 줄평균(기준행.get('SRED')) - 월급여)

    def 값식(ref, s, f, v):
        """수식과 그 값을 함께 적습니다."""
        칸들.append(f'<c r="{ref}" s="{s}">'
                    f'<f>{_xml(f)}</f><v>{float(v)!r}</v></c>')

    예상줄 = [('매출채권 회수', 'J', 매출식, 월매출),      # J = 차변(들어오는 돈)
              ('매입채무 지급', 'K', 원가식, 월원가),      # K = 대변(나가는 돈)
              ('급여', 'K', 급여식, 월급여),
              ('기타 비용', 'K', 기타식, 월기타)]
    for 이름, 금액열, 수식, 값 in 예상줄:
        for 월 in 예상월:
            r += 1
            숫(f'I{r}', S일자, _날짜일련(datetime.date(
                연도, 월, calendar.monthrange(연도, 월)[1])))
            for 열 in ('J', 'K'):
                if 열 == 금액열:
                    값식(f'{열}{r}', S숫자, f'ROUND({수식},0)', round(값, 0))
                else:
                    빈(f'{열}{r}', S숫자)
            글(f'L{r}', S가운데, '영업활동')
            글(f'M{r}', S왼쪽, 이름)
    오끝 = r + 1
    for 열 in 오열:
        빈(f'{열}{오끝}', S끝0 if 열 == 'I' else S끝)

    # ── 3. 예상 입금 ─────────────────────────────────────────
    #    지난 파일에 이미 있으면 담당자가 채운 그대로 옮깁니다 (한 칸도 안 건드립니다).
    입금항목 = [('영업활동', '매출채권 회수'), ('영업활동', '기타입금'),
                ('재무활동', '유상증자'), ('재무활동', '차입증감')]
    if 보존칸:
        입끝 = max(r for r, _c in 보존칸)
        for (행, 열), 조각 in sorted(보존칸.items()):
            칸들.append(조각)
    else:
        글('O3', S소제목, '3. 예상 현금성 자금 흐름(입금)')
        글('O4', S설명, f'- 담당자가 직접 채워 넣는 예상 입금 ({보고달 + 1}~12월)'
                        if 보고달 < 12 else '- 12월까지 실적이 다 들어와 예상분은 없습니다')
        for 열, 이름 in zip(입열, 예상머리):
            글(f'{열}5', S머리0 if 열 == 'O' else S머리, 이름)
        r = 6
        빈(f'O{r}', S일자0)
        빈(f'P{r}', S숫자)
        빈(f'Q{r}', S숫자)
        글(f'R{r}', S가운데, '기초')
        글(f'S{r}', S왼쪽, f'{연도}년 {보고달}월말 실제 잔액')
        for 구분이름, 이름 in 입금항목:
            for 월 in 예상월:
                r += 1
                숫(f'O{r}', S일자, _날짜일련(datetime.date(
                    연도, 월, calendar.monthrange(연도, 월)[1])))
                빈(f'P{r}', S숫자)           # ← 담당자가 여기에 금액을 넣습니다
                빈(f'Q{r}', S숫자)
                글(f'R{r}', S가운데, 구분이름)
                글(f'S{r}', S왼쪽, 이름)
        입끝 = r + 1
        for 열 in 입열:
            빈(f'{열}{입끝}', S끝0 if 열 == 'O' else S끝)

    # ── 작성 근거
    순 = (현['차변'] - 현['대변'])
    구분합 = (현.assign(순=순).groupby('항목')['순'].agg(['size', 'sum'])
              .sort_values('sum'))
    근거 = [
        ('■ 작성 근거', ''),
        ('', ''),
        ('[1] 실제 현금성 자금 흐름 (왼쪽)', ''),
        ('무엇을 옮겼나', f'{연도 % 100}년 원장_raw 에서 현금·예금 계정에 찍힌 줄만 날짜순으로 '
                         f'뽑았습니다 ({len(현):,}줄, {연도}년 1~{보고달}월). 현금 계정은 '
                         f'BS_IS_매핑에서 「현금」·「예금」이 들어간 재무상태표 계정으로 잡았습니다.'),
        ('차변 / 대변', '원장의 Debit / Credit 을 그대로 옮겼습니다. 손대지 않았습니다.'),
        ('기말', f'바로 윗줄 기말 + 차변 − 대변 (수식). 기초 {기초현금액:,.2f} 는 원장의 '
                 f'Beginning Balance 합계입니다.'),
        ('구분 · 활동성항목', '원장은 복식부기라, 현금이 든 전표에서 현금이 아닌 상대 계정이 그 돈의 '
                            '성격을 알려 줍니다. 전표(Transaction ID)가 같은 줄 중 현금이 아닌 것을 '
                            '찾아 그 계정으로 정했습니다. 상대가 여럿이면 금액이 가장 큰 것을 쓰고, '
                            '외환차손익은 곁다리라 뒤로 미뤘습니다.'),
        ('맞는지 확인', f'기초 {기초현금액:,.2f} + 순증감 {float(순.sum()):,.2f} '
                       f'= {기말액:,.2f} — 대시보드 현금흐름의 {보고달}월말 잔액과 같습니다.'),
        ('', ''),
        ('[2] 예상 현금성 자금 흐름 (오른쪽)  ※ 실적이 아니라 「가정」입니다', ''),
        ('가정 ①  매출액', f'월 {월매출:,.0f} — 1~{보고달}월 실제 매출을 {보고달}달로 나눈 '
                          f'월평균입니다. 매출채권 회수 줄에 그대로 넣었습니다 '
                          f'(그 달 매출은 그 달 안에 들어온다고 봅니다).'),
        ('가정 ②  매출원가', f'월 {월원가:,.0f} — 1~{보고달}월 실제 매출원가의 월평균입니다 '
                            f'(= 예상 매출액 × 실적 원가율). 매입채무 지급 줄에 넣었습니다.'),
        ('가정 ③  판관비(현금)', f'월 {월급여 + 월기타:,.0f} = 급여 {월급여:,.0f} + '
                                f'기타 비용 {월기타:,.0f}. 관리비 합계 월평균에서 '
                                f'상각비(현금이 나가지 않음)와 SR&ED 보조금(한 번 받고 마는 돈)을 '
                                f'뺀 금액입니다.'),
        ('수식으로 걸어 두었습니다', f"네 줄 모두 칸 안에 "
                                   f"=ROUND(AVERAGE('{집계시트}'!X…{끝열}…),0) 수식이 들어 있습니다. "
                                   f'다음 달 실적이 들어오면 예상액도 저절로 다시 계산됩니다 — '
                                   f'숫자를 손으로 고치실 필요가 없습니다.'),
        ('넣지 않은 것', '설비 투자 · 유상증자 · 차입은 회사가 정하는 일이라 여기 넣지 않았습니다. '
                       '운전자본(매출채권·매입채무·재고)도 지금 수준 그대로라고 봤습니다. '
                       '바꾸시려면 3번 표에 넣으시면 됩니다.'),
        ('기말', f'{보고달}월말 실제 잔액에서 출발해 한 줄씩 더하고 빼 내려갑니다.'),
        ('', ''),
        ('[3] 예상 현금성 자금 흐름(입금)', ''),
        ('무엇을 깔았나', '매출채권 회수 · 기타입금 · 유상증자 · 차입증감 네 항목을 '
                        f'{보고달 + 1}월부터 12월까지 줄로 깔아 두었습니다. '
                        '금액은 담당자가 「차변」 칸에 직접 넣으시면 됩니다.'
                        if 예상월 else '12월까지 실적이 다 들어와 예상분은 없습니다.'),
        ('넣으면 어떻게 되나', '대시보드 현금흐름 표의 「○월(예상)」 칸에 그대로 반영됩니다 — '
                            '출금(2번)과 입금(3번)을 합쳐 연말까지 잔액이 이어집니다.'),
    ]
    r2 = max(왼끝, 오끝, 입끝) + 2
    for 제목, 본문 in 근거:
        if 제목.startswith(('■', '[')):
            글(f'I{r2}', S근거제목, 제목)
        elif 제목:
            글(f'I{r2}', S근거, '· ' + 제목)
            글(f'K{r2}', S근거, 본문)
        r2 += 1
    if 예상월:
        글(f'I{r2}', S근거, '· 한 달 예상 순증감 (입금 − 출금, USD)')
        식(f'L{r2}', S근거숫자,
          f'(SUM($J$7:$J${오끝 - 1})-SUM($K$7:$K${오끝 - 1}))/{len(예상월)}')
        r2 += 2
    글(f'I{r2}', S근거제목,
      f'[3] 참고 — 실제 현금이 어디로 오갔나 ({연도}년 1~{보고달}월, USD)')
    r2 += 1
    for 이름, 줄 in 구분합.iterrows():
        글(f'I{r2}', S근거, '· ' + str(이름))
        글(f'K{r2}', S근거, f'{int(줄["size"]):,}줄')
        숫(f'L{r2}', S근거숫자, round(float(줄['sum']), 2))
        r2 += 1
    글(f'I{r2}', S근거, f'· 합계 (= {보고달}개월 순증감)')
    숫(f'L{r2}', S근거숫자, round(float(순.sum()), 2))

    폭 = ('<cols><col min="1" max="1" width="2.5" customWidth="1"/>'
          '<col min="2" max="2" width="13" customWidth="1"/>'
          '<col min="3" max="5" width="15.5" customWidth="1"/>'
          '<col min="6" max="6" width="14" customWidth="1"/>'
          '<col min="7" max="7" width="24" customWidth="1"/>'
          '<col min="8" max="8" width="3" customWidth="1"/>'
          '<col min="9" max="9" width="13" customWidth="1"/>'
          '<col min="10" max="11" width="15.5" customWidth="1"/>'
          '<col min="12" max="12" width="14" customWidth="1"/>'
          '<col min="13" max="13" width="24" customWidth="1"/>'
          '<col min="14" max="14" width="3" customWidth="1"/>'
          '<col min="15" max="15" width="13" customWidth="1"/>'
          '<col min="16" max="17" width="15.5" customWidth="1"/>'
          '<col min="18" max="18" width="14" customWidth="1"/>'
          '<col min="19" max="19" width="24" customWidth="1"/></cols>')
    행모음 = {}
    for c in 칸들:
        ref = re.match(r'<c r="([A-Z]+)(\d+)"', c)
        행모음.setdefault(int(ref.group(2)), []).append((ref.group(1), c))
    본문 = ''.join(
        f'<row r="{n}" spans="1:{끝열}">'
        + ''.join(x for _c, x in sorted(행모음[n], key=lambda t: (len(t[0]), t[0])))
        + '</row>' for n in sorted(행모음))
    끝행 = max(행모음)
    끝열 = max((_열번호(re.match(r'<c r="([A-Z]+)', x).group(1))
                for 줄 in 행모음.values() for _c, x in 줄), default=19)
    시트 = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetPr><tabColor theme="9" tint="0.79998168889431442"/></sheetPr>'
            f'<dimension ref="A1:{_열이름(끝열)}{끝행}"/>'
            '<sheetViews><sheetView showGridLines="0" workbookViewId="0">'
            '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="B6" sqref="B6"/></sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="17.4"/>'
            + 폭 + '<sheetData>' + 본문 + '</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
            'header="0.3" footer="0.3"/></worksheet>')
    안내 = {'시트': 자금시트, '현금줄': len(현), '기초': 기초현금액, '기말': 기말액,
            '예상월': 예상월, '항목수': len(예상줄), '서식추가': 새서식,
            '가정': {'월매출': 월매출, '월원가': 월원가,
                     '월급여': 월급여, '월기타': 월기타}}
    return 시트, (스타일 if 새서식 else None), 시트칸[자금시트], 안내


#  「당월 실적집계」에서 기준 월(H2) 대신 빈칸을 보고 있던 수식들.
#  Q18·R18 은 매출원가 누적인데 G5·H5(빈칸)를 참조해 OFFSET 이 0칸이 되어,
#  1월부터 이번 달까지가 아니라 「연간실적 칸 + 1월」을 더하고 있었습니다.
당월수식고침 = {
    'Q18': ("SUM(OFFSET('월별 실적집계'!$W$10,0,1):"
            "(OFFSET('월별 실적집계'!$W$10,0,G5)))",
            "SUM(OFFSET('월별 실적집계'!$W$10,0,1):"
            "(OFFSET('월별 실적집계'!$W$10,0,$H$2)))"),
    'R18': ("SUM(OFFSET('월별 실적집계'!$J$10,0,1):"
            "(OFFSET('월별 실적집계'!$J$10,0,H5)))",
            "SUM(OFFSET('월별 실적집계'!$J$10,0,1):"
            "(OFFSET('월별 실적집계'!$J$10,0,$H$2)))"),
}


def _당월수식고치기(시트xml):
    """잘못된 참조를 쓰던 칸의 수식을 바로잡습니다. 돌려주는 값 : (xml, 고친 칸 목록)"""
    고침 = []
    for 칸, (옛, 새) in 당월수식고침.items():
        def _바꿔(m, 칸=칸, 옛=옛, 새=새):
            속성, 몸 = m.group(1), m.group(2)
            fm = re.search(r'<f([^>]*)>(.*?)</f>', 몸, re.S)
            if not fm or _풀기(fm.group(2)).replace(' ', '') != 옛.replace(' ', ''):
                return m.group(0)
            고침.append(칸)
            return f'<c r="{칸}"{속성}><f{fm.group(1)}>{_xml(새)}</f></c>'
        시트xml = re.sub(rf'<c r="{칸}"([^>]*)>(.*?)</c>', _바꿔,
                         시트xml, count=1, flags=re.S)
    return 시트xml, 고침


# ══ 연결패키지 「검증」 시트 · 「퀵북 BS」「퀵북 IS」 시트 (2026-10-27) ══════════════
#   ① 원장 파일에 함께 올린 퀵북 BS·IS 를 패키지 안에 원본 그대로 넣고(「퀵북 BS」「퀵북 IS」),
#   BS 시트 앞 「검증」 시트가 패키지 BS·IS(당기 수정전)와 맞는지 수식으로 비교합니다.
#   · 양식(② 전월 패키지)에 두 시트가 이미 있으면 내용만 이번 달 것으로 갈아 끼웁니다(서식은 그대로).
#   · 없으면 세 시트를 새로 만들어 끼웁니다. 「검증」은 수식뿐이라 한 번 만든 뒤에는 손대지 않습니다.
검_퀵북시트 = {'BS': '퀵북 BS', 'IS': '퀵북 IS'}
검_검증시트 = '검증'
_검_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_검_시트형 = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'


def 검_엑스(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))


def 검_퀵북표읽기(원장바이트):
    """① 원장 파일의 퀵북 BS·IS 시트 → {'BS': [(A글, B값|'=수식'|None), ...], 'IS': [...]}.
       시트 이름이 「BS」「IS」 이거나 「9월(BS)」 처럼 (BS)·(IS) 가 붙은 것을 씁니다."""
    import openpyxl as _ox
    out = {}
    try:
        wb = _ox.load_workbook(io.BytesIO(원장바이트), read_only=True, data_only=False)
    except Exception:
        return out
    for 시트 in wb.sheetnames:
        이름 = str(시트).strip().upper()
        m = re.search(r'\(\s*(BS|IS)\s*\)\s*$', 이름)
        종류 = 이름 if 이름 in ('BS', 'IS') else (m.group(1) if m else None)
        if not 종류 or 종류 in out:
            continue
        줄들 = []
        for r in wb[시트].iter_rows(max_col=2, values_only=True):
            a = r[0] if len(r) > 0 else None
            b = r[1] if len(r) > 1 else None
            a = '' if a is None else str(a).strip()
            if isinstance(b, str) and not b.startswith('='):
                b = b.strip() or None
            줄들.append((a, b))
        while 줄들 and not 줄들[-1][0] and 줄들[-1][1] is None:
            줄들.pop()
        if any(a for a, _ in 줄들):
            out[종류] = 줄들
    wb.close()
    return out


def _검_계산(줄들):
    """「=B8+B9-B12」 같은 퀵북 합계 수식을 미리 계산해 둡니다 (엑셀 보호된 보기에서도 숫자가 보이게)."""
    값 = {}

    def 하나(i, 깊이=0):
        if i in 값:
            return 값[i]
        if i < 1 or i > len(줄들) or 깊이 > 400:
            return 0.0
        b = 줄들[i - 1][1]
        if isinstance(b, str) and b.startswith('='):
            식 = b[1:].replace(' ', '')
            if not re.fullmatch(r'[+-]?\$?B\$?\d+([+-]\$?B\$?\d+)*', 식):
                값[i] = None
                return None
            합 = 0.0
            for 부호, 행 in re.findall(r'([+-]?)\$?B\$?(\d+)', 식):
                v = 하나(int(행), 깊이 + 1) or 0.0
                합 += -v if 부호 == '-' else v
            값[i] = 합
        elif isinstance(b, (int, float)) and not isinstance(b, bool):
            값[i] = float(b)
        else:
            값[i] = None
        return 값[i]

    return [하나(i) for i in range(1, len(줄들) + 1)]


def 검_스타일더하기(sx):
    """검증 · 퀵북 시트용 서식을 styles.xml 맨 뒤에 덧붙입니다 → {역할: 번호}, 조건부 서식 dxf 번호."""
    번호 = 900
    while f'numFmtId="{번호}"' in sx:
        번호 += 1
    금액꼴 = f'<numFmt numFmtId="{번호}" formatCode="#,##0.00;[Red]\\-#,##0.00;&quot;-&quot;"/>'
    if '<numFmts' in sx:
        sx = re.sub(r'<numFmts count="(\d+)"', lambda m: f'<numFmts count="{int(m.group(1)) + 1}"', sx, count=1)
        sx = sx.replace('</numFmts>', 금액꼴 + '</numFmts>', 1)
    else:
        sx = sx.replace('<fonts', f'<numFmts count="1">{금액꼴}</numFmts><fonts', 1)

    m = re.search(r'<fonts count="(\d+)"[^>]*>(.*?)</fonts>', sx, re.S)
    글꼴수 = int(m.group(1))
    첫 = re.search(r'<font\b.*?</font>|<font/>', m.group(2), re.S).group(0)
    이름 = re.search(r'<name val="([^"]*)"', 첫)
    이름 = 이름.group(1) if 이름 else '맑은 고딕'

    def 글꼴(크기, 굵게=False, 색='FF000000'):
        return (f'<font>{"<b/>" if 굵게 else ""}<sz val="{크기}"/><color rgb="{색}"/>'
                f'<name val="{이름}"/><family val="3"/><charset val="129"/></font>')

    새글꼴 = [글꼴(10), 글꼴(10, True), 글꼴(14, True, 'FF1F3864'), 글꼴(10, True, 'FFFFFFFF'),
              글꼴(9, False, 'FF7F7F7F'), 글꼴(10, True, 'FFC00000'), 글꼴(12, True, 'FF1F3864')]
    sx = re.sub(r'<fonts count="(\d+)"', lambda mm: f'<fonts count="{int(mm.group(1)) + len(새글꼴)}"', sx, count=1)
    sx = sx.replace('</fonts>', ''.join(새글꼴) + '</fonts>', 1)
    F = {k: 글꼴수 + i for i, k in enumerate(['보통', '굵게', '제목', '흰굵게', '회색', '빨강', '소제목'])}

    m = re.search(r'<fills count="(\d+)"', sx)
    채움수 = int(m.group(1))
    새채움 = ['<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>',
              '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>']
    sx = re.sub(r'<fills count="(\d+)"', lambda mm: f'<fills count="{int(mm.group(1)) + len(새채움)}"', sx, count=1)
    sx = sx.replace('</fills>', ''.join(새채움) + '</fills>', 1)
    남색, 회색칸 = 채움수, 채움수 + 1

    m = re.search(r'<borders count="(\d+)"', sx)
    밑줄 = int(m.group(1))
    sx = re.sub(r'<borders count="(\d+)"', lambda mm: f'<borders count="{int(mm.group(1)) + 1}"', sx, count=1)
    sx = sx.replace('</borders>', '<border><left/><right/><top/><bottom style="thin"><color rgb="FFBFBFBF"/>'
                    '</bottom><diagonal/></border></borders>', 1)

    def xf(글, 꼴=0, 채=0, 테=0, 가로=None, 줄바꿈=False):
        정렬 = ''
        if 가로 or 줄바꿈:
            정렬 = ('<alignment' + (f' horizontal="{가로}"' if 가로 else '') + ' vertical="center"'
                    + (' wrapText="1"' if 줄바꿈 else '') + '/>')
        return (f'<xf numFmtId="{꼴}" fontId="{글}" fillId="{채}" borderId="{테}" xfId="0"'
                + (' applyNumberFormat="1"' if 꼴 else '') + ' applyFont="1"'
                + (' applyFill="1"' if 채 else '') + (' applyBorder="1"' if 테 else '')
                + (' applyAlignment="1">' + 정렬 + '</xf>' if 정렬 else '/>'))

    역할 = [
        ('제목', xf(F['제목'])),
        ('설명', xf(F['회색'])),
        ('소제목', xf(F['소제목'])),
        ('머리', xf(F['흰굵게'], 0, 남색, 0, 'center', True)),
        ('글자', xf(F['보통'], 0, 0, 밑줄)),
        ('숫자', xf(F['보통'], 번호, 0, 밑줄, 'right')),
        ('굵은글자', xf(F['굵게'], 0, 0, 밑줄)),
        ('굵은숫자', xf(F['굵게'], 번호, 0, 밑줄, 'right')),
        ('판정', xf(F['굵게'], 0, 0, 밑줄, 'center')),
        ('결과', xf(F['소제목'], 0, 회색칸, 0, 'left')),
        ('빨강', xf(F['빨강'])),
        ('퀵북제목', xf(F['굵게'])),
        ('퀵북글자', xf(F['보통'])),
        ('퀵북숫자', xf(F['보통'], 번호, 0, 0, 'right')),
        ('퀵북굵은숫자', xf(F['굵게'], 번호, 0, 0, 'right')),
    ]
    m = re.search(r'<cellXfs count="(\d+)"', sx)
    첫칸 = int(m.group(1))
    sx = re.sub(r'<cellXfs count="(\d+)"', lambda mm: f'<cellXfs count="{int(mm.group(1)) + len(역할)}"', sx, count=1)
    sx = sx.replace('</cellXfs>', ''.join(x for _, x in 역할) + '</cellXfs>', 1)
    서식 = {k: 첫칸 + i for i, (k, _) in enumerate(역할)}

    # 조건부 서식 : 「확인 필요」 빨강 바탕 · 「일치」 초록 글자
    dxf = ('<dxf><font><b/><color rgb="FFC00000"/></font><fill><patternFill><bgColor rgb="FFFCE4E4"/>'
           '</patternFill></fill></dxf>'
           '<dxf><font><b/><color rgb="FF00804A"/></font></dxf>')
    if re.search(r'<dxfs count="(\d+)"\s*/>', sx):
        sx = re.sub(r'<dxfs count="\d+"\s*/>', f'<dxfs count="2">{dxf}</dxfs>', sx, count=1)
        첫dxf = 0
    elif '<dxfs' in sx:
        첫dxf = int(re.search(r'<dxfs count="(\d+)"', sx).group(1))
        sx = re.sub(r'<dxfs count="(\d+)"', lambda mm: f'<dxfs count="{int(mm.group(1)) + 2}"', sx, count=1)
        sx = sx.replace('</dxfs>', dxf + '</dxfs>', 1)
    else:
        첫dxf = 0
        sx = sx.replace('</cellStyles>', f'</cellStyles><dxfs count="2">{dxf}</dxfs>', 1) \
            if '</cellStyles>' in sx else sx.replace('</cellXfs>', f'</cellXfs><dxfs count="2">{dxf}</dxfs>', 1)
    서식['dxf빨강'], 서식['dxf초록'] = 첫dxf, 첫dxf + 1
    return sx, 서식


def 검_퀵북값(퀵북):
    """퀵북 BS·IS 줄 → 검증 시트 B열 이름별 값 (없으면 None)."""
    out = {}
    for 종류 in ('BS', 'IS'):
        줄들 = 퀵북.get(종류) or []
        값 = _검_계산(줄들) if 줄들 else []
        표 = {}
        for (a, _b), v in zip(줄들, 값):
            if a and v is not None:
                표.setdefault(re.sub(r'\s+', ' ', a).strip().lower(), v)
        머리 = 줄들[2][0] if len(줄들) > 2 else ''
        꼬리 = next((a for a, b in reversed(줄들) if a), '') if 줄들 else ''
        if 종류 == 'BS':
            out['퀵북 BS 기준일'] = 머리 or None
            out['퀵북 내려받은 시각'] = 꼬리 or None
            for k, q in (('자산총계', 'total assets'), ('부채총계', 'total liabilities'),
                         ('자본총계', 'total equity'), ('부채및자본총계', 'total liabilities and equity')):
                out[k] = 표.get(q)
            out['GST'] = 표.get('gst/hst payable', 0.0) if 표 else None
        else:
            out['퀵북 IS 기간'] = 머리 or None
            for k, q in (('매출액', 'total income'), ('매출원가', 'total cost of goods sold'),
                         ('매출총이익', 'gross profit'), ('판매비와관리비', 'total expenses'),
                         ('당기순이익', 'profit')):
                out[k] = 표.get(q)
            oi, oe = 표.get('total other income'), 표.get('total other expenses')
            out['영업외손익 (순액)'] = (None if oi is None and oe is None
                                       else (oi or 0.0) - (oe or 0.0))
    return out


검_재분류줄 = ('자산총계', '부채총계', '부채및자본총계')


def 검_검증시트XML(서식, 줄번호, 퀵값):
    """「검증」 시트. 패키지 쪽은 수식, 퀵북 쪽은 ① 원장 파일에서 읽은 값(상수)입니다.
       줄번호 = {'BS': {'자산총계': 93, ...}, 'IS': {...}}"""
    R = "'26년 원장_raw'"
    _달 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

    def 달식(x):      # 글 안에 나오는 가장 늦은 달 (「January-September」 → 9, 「As of Sep 30」 → 9)
        식 = '0'
        for i, m in enumerate(_달, 1):
            식 = f'IF(ISNUMBER(SEARCH("{m}",{x})),{i},{식})'
        return 식
    B, I = 줄번호['BS'], 줄번호['IS']
    행들 = {}

    def 칸(ref, 내용, 역할, 식=False):
        s = 서식[역할]
        if 식:
            return f'<c r="{ref}" s="{s}"><f>{검_엑스(내용)}</f></c>'
        if isinstance(내용, (int, float)):
            return f'<c r="{ref}" s="{s}"><v>{내용}</v></c>'
        return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{검_엑스(내용)}</t></is></c>'

    def 넣기(행, *칸들):
        행들.setdefault(행, []).extend(칸들)

    넣기(2, 칸('B2', '검증 — 업로드한 퀵북 BS·IS 와 연결패키지 일치 여부', '제목'))
    넣기(3, 칸('B3', '패키지 BS·IS 「당기 (26년-수정전)」 I열 ↔ ① 원장 파일에 함께 올린 퀵북 BS·IS 의 금액 '
               '(대시보드가 생성할 때 적어 넣음). 수정분개(J·K열)는 비교하지 않습니다.', '설명'))
    넣기(5, 칸('B5', '결과', '소제목'),
         칸('C5', 'IF(COUNTIF($I$9:$I$40,"확인 필요")=0,"✔ 모두 일치","✘ 확인 필요 "&COUNTIF($I$9:$I$40,"확인 필요")&"건")',
            '결과', True))

    # ① 업로드 확인
    넣기(7, 칸('B7', '① 업로드 확인', '소제목'))
    넣기(8, *[칸(f'{c}8', t, '머리') for c, t in
               zip('BCDEFGHI', ['항목', '업로드 자료', '', '', '', '', '달', '판정'])])
    월원장 = f'IF(COUNT({R}!$C:$C)=0,0,MONTH(MAX({R}!$C:$C)))'
    줄 = [
        (9, '원장 기간 (26년 원장_raw)', f'IF(COUNT({R}!$C:$C)=0,"없음",{R}!$A$3)', 월원장,
         f'IF(COUNT({R}!$C:$C)=0,"확인 필요","기준")'),
        (10, '퀵북 BS 기준일', 퀵값.get('퀵북 BS 기준일') or '없음 — ① 원장 파일에 BS 시트가 없었습니다',
         달식('C10'), f'IF(H10=H$9,"일치","확인 필요")'),
        (11, '퀵북 IS 기간', 퀵값.get('퀵북 IS 기간') or '없음 — ① 원장 파일에 IS 시트가 없었습니다',
         달식('C11'), f'IF(H11=H$9,"일치","확인 필요")'),
        (12, '패키지 BS 기준일', 'BS!$B$2',
         'IFERROR(VALUE(MID(BS!$B$2,FIND("년",BS!$B$2)+2,2)),0)', f'IF(H12=H$9,"일치","확인 필요")'),
        (13, '원장 거래 줄 수', f'TEXT(COUNT({R}!$C:$C),"#,##0")&" 줄"', None, None),
        (14, '퀵북 내려받은 시각', 퀵값.get('퀵북 내려받은 시각') or '', None, None),
    ]
    상수칸 = {'퀵북 BS 기준일', '퀵북 IS 기간', '퀵북 내려받은 시각'}
    for 행, 이름, 내용, 달, 판정 in 줄:
        넣기(행, 칸(f'B{행}', 이름, '굵은글자'), 칸(f'C{행}', 내용, '글자', 이름 not in 상수칸),
             *[칸(f'{c}{행}', '', '글자') for c in 'DEFG'])
        넣기(행, 칸(f'H{행}', 달, '판정', True) if 달 else 칸(f'H{행}', '', '글자'),
             칸(f'I{행}', 판정, '판정', True) if 판정 else 칸(f'I{행}', '', '글자'))

    # ② BS · ③ IS
    머리 = ['과목', '패키지 (수정전)', '퀵북', '차이', '기초 차이', '재분류', '남는 차이', '판정']

    def 표(시작, 제목, 줄들):
        넣기(시작, 칸(f'B{시작}', 제목, '소제목'))
        넣기(시작 + 1, *[칸(f'{c}{시작 + 1}', t, '머리') for c, t in zip('BCDEFGHI', 머리)])
        r = 시작 + 2
        for 이름, 패식, 기초식 in 줄들:
            퀵 = 퀵값.get(이름)
            재분 = -(퀵값.get('GST') or 0.0) if 이름 in 검_재분류줄 else 0
            넣기(r, 칸(f'B{r}', 이름, '굵은글자'), 칸(f'C{r}', 패식, '숫자', True),
                 칸(f'D{r}', round(퀵, 2), '숫자') if 퀵 is not None else f'<c r="D{r}" s="{서식["숫자"]}"/>',
                 칸(f'E{r}', f'IF(D{r}="","",C{r}-D{r})', '숫자', True),
                 칸(f'F{r}', 기초식, '숫자', True) if 기초식 else 칸(f'F{r}', 0, '숫자'),
                 칸(f'G{r}', round(재분, 2) + 0.0, '숫자'),
                 칸(f'H{r}', f'IF(D{r}="","",ROUND(E{r}-F{r}-G{r},2))', '숫자', True),
                 칸(f'I{r}', f'IF(D{r}="","확인 필요",IF(ABS(H{r})<0.01,"일치","확인 필요"))', '판정', True))
            r += 1
        return r

    끝 = 표(16, '② 재무상태표 (BS)', [
        ('자산총계', f'BS!$I${B["자산총계"]}', None),
        ('부채총계', f'BS!$I${B["부채총계"]}', "-SUM('26년 BS기초'!$L$76:$L$113)"),
        ('자본총계', f'BS!$I${B["자본총계"]}', '-F19'),
        ('부채및자본총계', f'BS!$I${B["부채및자본총계"]}', None),
    ])
    끝 = 표(끝 + 1, '③ 손익계산서 (IS, 1월~당월 누적)', [
        ('매출액', f'IS!$I${I["매출액"]}', None),
        ('매출원가', f'IS!$I${I["매출원가"]}', None),
        ('매출총이익', f'IS!$I${I["매출총이익"]}', None),
        ('판매비와관리비', f'IS!$I${I["판매비와관리비"]}', None),
        ('영업외손익 (순액)',
         f'IS!$I${I["기타수익"]}+IS!$I${I["금융수익"]}-IS!$I${I["기타비용"]}-IS!$I${I["금융비용"]}', None),
        ('당기순이익', f'IS!$I${I["당기순이익"]}', None),
    ])
    넣기(끝 + 1, 칸(f'B{끝 + 1}', '※ 기초 차이 = 「26년 BS기초」 부채 계정의 (본사 확정 기초 − 퀵북 기초) — 지금은 '
                    'Interest payables (본사 이자비용 2024·2025 기초조정). 자본총계는 그 반대 금액.', '설명'))
    넣기(끝 + 2, 칸(f'B{끝 + 2}', '※ 재분류 = 퀵북 부채의 GST/HST Payable 을 패키지는 자산 「부가세대급금」 으로 표시.',
                    '설명'))
    끝 += 1
    xml행 = ''.join(f'<row r="{r}"' + (' ht="24" customHeight="1"' if r in (2, 5) else '') + '>'
                    + ''.join(sorted(c, key=lambda x: (len(re.search(r'r="([A-Z]+)', x).group(1)),
                                                       re.search(r'r="([A-Z]+)', x).group(1))))
                    + '</row>' for r, c in sorted(행들.items()))
    def 조건(범위, 첫칸, 우선):
        return (f'<conditionalFormatting sqref="{범위}">'
                f'<cfRule type="containsText" dxfId="{서식["dxf빨강"]}" priority="{우선}" operator="containsText" '
                f'text="확인 필요"><formula>NOT(ISERROR(SEARCH("확인 필요",{첫칸})))</formula></cfRule>'
                f'<cfRule type="containsText" dxfId="{서식["dxf초록"]}" priority="{우선 + 1}" operator="containsText" '
                f'text="일치"><formula>NOT(ISERROR(SEARCH("일치",{첫칸})))</formula></cfRule>'
                '</conditionalFormatting>')
    cf = 조건('C5', 'C5', 1) + 조건('I9:I40', 'I9', 3)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{_검_NS}">'
            '<sheetPr><tabColor rgb="FFC00000"/></sheetPr>'
            f'<dimension ref="B2:I{끝 + 1}"/>'
            '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="18"/>'
            '<cols><col min="1" max="1" width="2" customWidth="1"/>'
            '<col min="2" max="2" width="26" customWidth="1"/>'
            '<col min="3" max="8" width="16" customWidth="1"/>'
            '<col min="9" max="9" width="12" customWidth="1"/></cols>'
            '<sheetData>' + xml행 + '</sheetData>'
            '<mergeCells count="6"><mergeCell ref="C9:G9"/><mergeCell ref="C10:G10"/>'
            '<mergeCell ref="C11:G11"/><mergeCell ref="C12:G12"/><mergeCell ref="C13:G13"/>'
            '<mergeCell ref="C14:G14"/></mergeCells>'
            + cf +
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            '</worksheet>')


def 검_줄번호(z, 지도, 공유):
    """패키지 BS·IS 에서 비교할 줄 번호를 국문 과목(E열)으로 찾습니다."""
    찾을것 = {'BS': ['자산총계', '부채총계', '자본총계', '부채및자본총계'],
              'IS': ['매출액', '매출원가', '매출총이익', '판매비와관리비', '기타수익', '기타비용',
                     '금융수익', '금융비용', '당기순이익(손실)']}
    out = {}
    for 시트, 목록 in 찾을것.items():
        xml = z.read(지도[시트]).decode('utf-8')
        찾음 = {}
        for m in re.finditer(r'<c r="E(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S):
            속, 내용 = m.group(2), m.group(3) or ''
            if 't="s"' in 속:
                v = re.search(r'<v>(\d+)</v>', 내용)
                글 = 공유[int(v.group(1))] if v else ''
            else:
                글 = _html.unescape(re.sub(r'<[^>]+>', '', 내용))
            글 = re.sub(r'\s+', '', str(글))
            if 글 in 목록 and 글 not in 찾음:
                찾음[글] = int(m.group(1))
        빠짐 = [k for k in 목록 if k not in 찾음]
        if 빠짐:
            raise ValueError(f'{시트} 시트에서 {", ".join(빠짐)} 줄을 찾지 못했습니다')
        찾음['당기순이익'] = 찾음.pop('당기순이익(손실)', 0) if 시트 == 'IS' else None
        out[시트] = {k: v for k, v in 찾음.items() if v is not None}
    return out


def 검_시트끼우기(파일바이트, 새시트들, 스타일xml=None):
    """새시트들 = [(이름, xml, 앞에올시트이름 | None=맨끝)] → 새 파일 바이트.
       workbook.xml · rels · [Content_Types] · 정의된이름(localSheetId)·activeTab 을 함께 고칩니다."""
    z = zipfile.ZipFile(io.BytesIO(파일바이트))
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rel = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    ct = z.read('[Content_Types].xml').decode('utf-8')
    있는파일 = set(z.namelist())
    번호 = 1
    새파트 = {}
    for 이름, xml, 앞 in 새시트들:
        while f'xl/worksheets/sheet{번호}.xml' in 있는파일 or f'xl/worksheets/sheet{번호}.xml' in 새파트:
            번호 += 1
        파일 = f'xl/worksheets/sheet{번호}.xml'
        새파트[파일] = xml
        rid = 1
        while re.search(rf'\bId="rIdQ{rid}"', rel):
            rid += 1
        rid = f'rIdQ{rid}'
        rel = rel.replace('</Relationships>', f'<Relationship Id="{rid}" Type="{_검_시트형}" '
                          f'Target="worksheets/sheet{번호}.xml"/></Relationships>', 1)
        ct = ct.replace('</Types>', f'<Override PartName="/{파일}" ContentType="application/'
                        'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>', 1)
        ids = [int(x) for x in re.findall(r'<sheet\b[^>]*\bsheetId="(\d+)"', wb)]
        새태그 = f'<sheet name="{검_엑스(이름)}" sheetId="{max(ids or [0]) + 1}" r:id="{rid}"/>'
        시트태그 = re.findall(r'<sheet\b[^>]*/>', wb)
        자리 = len(시트태그)
        for i, t in enumerate(시트태그):
            n = _html.unescape(re.search(r'\bname="([^"]*)"', t).group(1))
            if 앞 is not None and n == 앞:
                자리 = i
                break
        if 자리 < len(시트태그):
            wb = wb.replace(시트태그[자리], 새태그 + 시트태그[자리], 1)
        else:
            wb = wb.replace('</sheets>', 새태그 + '</sheets>', 1)
        wb = re.sub(r'localSheetId="(\d+)"',
                    lambda m: f'localSheetId="{int(m.group(1)) + (1 if int(m.group(1)) >= 자리 else 0)}"', wb)
        wb = re.sub(r'activeTab="(\d+)"',
                    lambda m: f'activeTab="{int(m.group(1)) + (1 if int(m.group(1)) >= 자리 else 0)}"', wb)
        wb = re.sub(r'firstSheet="(\d+)"',
                    lambda m: f'firstSheet="{int(m.group(1)) + (1 if int(m.group(1)) >= 자리 else 0)}"', wb)
    wb = re.sub(r'<calcPr[^>]*/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wb, count=1)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as out:
        for it in z.infolist():
            if it.filename == 'xl/calcChain.xml':
                continue
            d = z.read(it.filename)
            if it.filename == 'xl/workbook.xml':
                d = wb.encode('utf-8')
            elif it.filename == 'xl/_rels/workbook.xml.rels':
                d = re.sub(r'<Relationship[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', '', rel).encode('utf-8')
            elif it.filename == '[Content_Types].xml':
                d = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '', ct).encode('utf-8')
            elif it.filename == 'xl/styles.xml' and 스타일xml is not None:
                d = 스타일xml.encode('utf-8')
            out.writestr(it, d)
        for 파일, xml in 새파트.items():
            out.writestr(파일, xml.encode('utf-8'))
    return buf.getvalue()


def 검_값지우기(xml):
    """수식 칸에 남은 지난달 계산값을 지웁니다 — 엑셀 「보호된 보기」에서 옛 숫자로 「일치」가 보이지 않게
       (편집 사용을 누르면 엑셀이 새로 계산합니다)."""
    def _한칸(m):
        속, 몸 = m.group(1), m.group(2)
        if '<f' not in 몸:
            return m.group(0)
        속 = re.sub(r'\s+t="[^"]*"', '', 속)
        return f'<c{속}>' + re.sub(r'<v>.*?</v>|<v/>', '', 몸, flags=re.S) + '</c>'
    return re.sub(r'<c\b([^>]*?)(?<!/)>(.*?)</c>', _한칸, xml, flags=re.S)


def 검_시트지우기(파일바이트, 이름):
    """시트 하나를 통째로 뺍니다 (없으면 그대로)."""
    z = zipfile.ZipFile(io.BytesIO(파일바이트))
    지울, 패치 = _시트빼기(z, 이름)
    if not 지울:
        return 파일바이트
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as out:
        for it in z.infolist():
            if it.filename in 지울:
                continue
            out.writestr(it, 패치.get(it.filename) or z.read(it.filename))
    return buf.getvalue()


def 검_칸쓰기(xml, 칸, 값):
    """칸 하나의 값을 바꿉니다 (서식 번호는 그대로). 값 None 이면 빈칸."""
    def _바꿈(m):
        sty = re.search(r'\ss="(\d+)"', m.group(1) or '')
        sty = f' s="{sty.group(1)}"' if sty else ''
        if 값 is None or 값 == '':
            return f'<c r="{칸}"{sty}/>'
        if isinstance(값, (int, float)):
            return f'<c r="{칸}"{sty}><v>{float(값)!r}</v></c>'
        return (f'<c r="{칸}"{sty} t="inlineStr"><is><t xml:space="preserve">'
                f'{검_엑스(값)}</t></is></c>')
    return re.sub(rf'<c r="{칸}"((?:\s[^>]*?)?)(?:/>|(?<!/)>.*?</c>)', _바꿈, xml, count=1, flags=re.S)


def 검_검증갱신(xml, 퀵값, 공유):
    """이미 있는 「검증」 시트 — B열 항목 이름을 보고 퀵북 칸(C·D·G)만 이번 달 값으로 바꿉니다."""
    상수C = {'퀵북 BS 기준일': '없음 — ① 원장 파일에 BS 시트가 없었습니다',
             '퀵북 IS 기간': '없음 — ① 원장 파일에 IS 시트가 없었습니다', '퀵북 내려받은 시각': ''}
    숫자줄 = {'자산총계', '부채총계', '자본총계', '부채및자본총계', '매출액', '매출원가', '매출총이익',
              '판매비와관리비', '영업외손익 (순액)', '당기순이익'}
    for m in list(re.finditer(r'<c r="B(\d+)"([^>]*?)(?:/>|(?<!/)>(.*?)</c>)', xml, re.S)):
        행, 속, 몸 = m.group(1), m.group(2), m.group(3) or ''
        if 't="s"' in 속:
            v = re.search(r'<v>(\d+)</v>', 몸)
            글 = 공유[int(v.group(1))] if v and int(v.group(1)) < len(공유) else ''
        else:
            글 = _html.unescape(re.sub(r'<[^>]+>', '', 몸))
        글 = str(글).strip()
        if 글 in 상수C:
            xml = 검_칸쓰기(xml, f'C{행}', 퀵값.get(글) or 상수C[글])
        elif 글 in 숫자줄:
            v = 퀵값.get(글)
            xml = 검_칸쓰기(xml, f'D{행}', None if v is None else round(v, 2))
            xml = 검_칸쓰기(xml, f'G{행}', round(-(퀵값.get('GST') or 0.0), 2) + 0.0
                            if 글 in 검_재분류줄 else 0.0)
    return 검_값지우기(xml)


과제시트들 = ('조직운영안',)


def 과제잠금(xml, 보고, 공유):
    """「개선과제」「판관비절감」 시트 — 지난 달 칸은 잠그고(시트 보호), 당월 · 앞으로의 달 칸만 고칠 수 있게 합니다.
       보고 = (연, 월). 칸 제목 「2026년 9월 — …」 을 B열에서 찾아 칸 범위를 정합니다."""
    제목들 = []
    for m in re.finditer(r'<c r="B(\d+)"([^>]*?)(?:/>|(?<!/)>(.*?)</c>)', xml, re.S):
        속, 몸 = m.group(2), m.group(3) or ''
        if 't="s"' in 속:
            v = re.search(r'<v>(\d+)</v>', 몸)
            글 = 공유[int(v.group(1))] if v and int(v.group(1)) < len(공유) else ''
        else:
            글 = _html.unescape(re.sub(r'<[^>]+>', '', 몸))
        mm = re.match(r'^\s*(20\d\d)\s*년\s*(\d{1,2})\s*월\s*—', str(글))
        if mm:
            제목들.append((int(m.group(1)), (int(mm.group(1)), int(mm.group(2)))))
    제목들.sort()
    if not 제목들:
        return xml
    끝행 = max([int(x) for x in re.findall(r'<row r="(\d+)"', xml)] or [0])
    열린 = []
    for i, (행, ym) in enumerate(제목들):
        끝 = (제목들[i + 1][0] - 2) if i + 1 < len(제목들) else 끝행
        if ym >= tuple(보고) and 끝 >= 행 + 2:
            열린.append(f'B{행 + 2}:I{끝}')
    xml = re.sub(r'<sheetProtection\b[^>]*/>', '', xml)
    xml = re.sub(r'<protectedRanges>.*?</protectedRanges>', '', xml, flags=re.S)
    보호 = ('<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0" formatRows="0" '
            'insertRows="0" deleteRows="1" selectLockedCells="0" selectUnlockedCells="0"/>'
            + (f'<protectedRanges><protectedRange sqref="{" ".join(열린)}" name="당월이후"/></protectedRanges>' if 열린 else ''))
    return xml.replace('</sheetData>', '</sheetData>' + 보호, 1) if '</sheetData>' in xml else xml


def 과제잠그기(파일바이트, 보고):
    """연결패키지의 과제 시트들에 이번 달 기준 잠금을 다시 겁니다 (매월 생성할 때)."""
    z = zipfile.ZipFile(io.BytesIO(파일바이트))
    지도 = dict(_시트지도(z))
    대상 = {지도[n]: n for n in 과제시트들 if 지도.get(n)}
    if not 대상:
        return 파일바이트
    공유 = 패_공유(z)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as out:
        for it in z.infolist():
            d = z.read(it.filename)
            if it.filename in 대상:
                d = 과제잠금(d.decode('utf-8'), 보고, 공유).encode('utf-8')
            out.writestr(it, d)
    return buf.getvalue()


def 검_패키지에넣기(파일바이트, 퀵북):
    """연결패키지 파일의 「검증」 시트(BS 앞)에 이번 달 퀵북 BS·IS 금액을 적어 넣습니다.
       처음이면 시트를 새로 만들고, 있으면 퀵북 칸만 바꿉니다(서식은 담당자 것 그대로).
       ※ 퀵북 원본 시트는 패키지에 넣지 않습니다 — 대시보드에 마지막 원장 파일로 남아 있습니다 (2026-10-27)
       돌려주는 값 : (새 파일 바이트, 안내 글)"""
    지도 = dict(_시트지도(zipfile.ZipFile(io.BytesIO(파일바이트))))
    if not (지도.get('BS') and 지도.get('IS')):
        return 파일바이트, ''
    for 이름 in 검_퀵북시트.values():             # 예전(v8)에 넣었던 퀵북 원본 시트는 뺍니다
        if 이름 in 지도:
            파일바이트 = 검_시트지우기(파일바이트, 이름)
    z = zipfile.ZipFile(io.BytesIO(파일바이트))
    지도 = dict(_시트지도(z))
    퀵값 = 검_퀵북값(퀵북)
    안내 = ('퀵북 BS·IS 금액 반영' if 퀵북.get('BS') and 퀵북.get('IS')
            else '퀵북 ' + '·'.join(k for k in ('BS', 'IS') if not 퀵북.get(k)) + ' 없음 (원장 파일에 시트 없음)')
    if 지도.get(검_검증시트):
        옛 = z.read(지도[검_검증시트]).decode('utf-8')
        if "퀵북 BS'!" in 옛 or '퀵북 BS&apos;!' in 옛:     # 예전 모양(퀵북 시트를 보는 수식) → 새로 만듦
            파일바이트 = 검_시트지우기(파일바이트, 검_검증시트)
            z = zipfile.ZipFile(io.BytesIO(파일바이트))
            지도 = dict(_시트지도(z))
        else:
            새 = 검_검증갱신(옛, 퀵값, 패_공유(z))
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as out:
                for it in z.infolist():
                    out.writestr(it, 새.encode('utf-8') if it.filename == 지도[검_검증시트]
                                 else z.read(it.filename))
            return buf.getvalue(), 안내
    sx, 서식 = 검_스타일더하기(z.read('xl/styles.xml').decode('utf-8'))
    줄번호 = 검_줄번호(z, {k: 지도[k] for k in ('BS', 'IS')}, 패_공유(z))
    return 검_시트끼우기(파일바이트, [(검_검증시트, 검_검증시트XML(서식, 줄번호, 퀵값), 'BS')], sx), 안내


import copy
import openpyxl
from openpyxl.formula.translate import Translator
import glob


def _비밀값(열쇠, 기본):
    """거래처 이름 · 대출 숫자 같은 회사 정보는 코드(GitHub)에 두지 않고 Streamlit 설정(Secrets)에서 읽습니다."""
    try:
        v = st.secrets.get(열쇠)
    except Exception:
        v = None
    return 기본 if v is None else v


커_고객별칭 = {k: list(v) for k, v in dict(_비밀값('ctk_customer_alias', {})).items()}   # 통장 적요에서 거래처를 찾을 글자 (Secrets)
커_거래처별칭 = {k: list(v) for k, v in dict(_비밀값('ctk_vendor_alias', {})).items()}   # 지급처 (Secrets)
커_팩토링거래처 = [str(x) for x in _비밀값('ctk_factoring_customers', [])]   # 채권채무 ① 에서 원장처럼 건별로 보여 줄 팩토링 거래처 (Secrets · 나머지는 집계)


def 커_이름정리(n):
    n = '' if n is None or (isinstance(n, float) and pd.isna(n)) else str(n).strip()
    return re.sub(r'\s*\((IC\s+)?(Customer|Customers|Vendor)\)\s*$', '', n, flags=re.I).strip()


def 커_원장읽기(바이트):
    """CTKonnect 퀵북 원장 → 표 (계정 · 일자 · 거래처 · 적요 · 차변 · 대변 …)."""
    xls = pd.ExcelFile(io.BytesIO(바이트))
    d = None
    for s in xls.sheet_names:
        for h in (0, 1, 2, 3, 4, 5):
            try:
                t = pd.read_excel(xls, sheet_name=s, header=h)
            except Exception:
                continue
            if 'Distribution account' in [str(c) for c in t.columns] and 'Transaction date' in [str(c) for c in t.columns]:
                d = t
                break
        if d is not None:
            break
    if d is None:
        raise ValueError('원장 파일에서 「Distribution account · Transaction date」 열을 찾지 못했습니다.')
    날 = pd.to_datetime(d['Transaction date'], format='%m/%d/%Y', errors='coerce')
    날 = 날.fillna(pd.to_datetime(d['Transaction date'], errors='coerce'))
    d = d.assign(일자=날)
    d = d[d['일자'].notna() & d['Distribution account'].notna()].copy()
    d['계정'] = d['Distribution account'].astype(str).str.strip()
    d = d[~d['계정'].str.startswith(('Total', 'Beginning'))]
    d['차'] = pd.to_numeric(d.get('Debit'), errors='coerce').fillna(0.0)
    d['대'] = pd.to_numeric(d.get('Credit'), errors='coerce').fillna(0.0)
    v = d['Vendor'] if 'Vendor' in d.columns else pd.Series([None] * len(d), index=d.index)
    c = d['Customer'] if 'Customer' in d.columns else pd.Series([None] * len(d), index=d.index)
    if 'Vendor' not in d.columns and 'Customer' not in d.columns and 'Name' in d.columns:
        v = c = d['Name']           # 「General Ledger List」 양식 — 거래처가 Name 한 열 (분개 줄은 비어 있는 경우가 많음)
    d['거래처원'] = v.where(v.notna(), c)
    d['거래처원'] = d['거래처원'].where(d['거래처원'].notna(), None)
    d['고객'] = c.where(c.notna(), None)
    d['공급'] = v.where(v.notna(), None)
    memo = d['Memo'] if 'Memo' in d.columns else d.get('Description')
    d['적요'] = memo.fillna('').astype(str).str.replace('\n', ' ') if memo is not None else ''
    d['TID'] = d.get('Transaction ID')
    d['Num'] = d['Num'] if 'Num' in d.columns else None
    return d.reset_index(drop=True)


def 커_통장읽기(바이트):
    """2026년 금융거래 파일의 「자금data(US)」 → 일자 · 적요 · 입금 · 출금 · 세부."""
    if not 바이트:
        return None
    xls = pd.ExcelFile(io.BytesIO(바이트))
    시트 = next((s for s in xls.sheet_names if '자금data' in s and 'US' in s), None)
    if 시트 is None:
        raise ValueError('통장 파일에 「자금data(US)」 시트가 없습니다 — 「2026년 금융거래」 파일을 올려 주세요.')
    raw = pd.read_excel(xls, sheet_name=시트, header=None)
    머리 = next(i for i in range(min(len(raw), 20)) if '거래일자' in [str(x).strip() for x in raw.iloc[i].tolist()])
    d = pd.read_excel(xls, sheet_name=시트, header=머리)
    d.columns = [str(c).strip() for c in d.columns]
    d['일자'] = pd.to_datetime(d['거래일자'], errors='coerce')
    d = d[d['일자'].notna()].copy()
    d['입금'] = pd.to_numeric(d.get('입금액'), errors='coerce').fillna(0.0)
    d['출금'] = pd.to_numeric(d.get('출금액'), errors='coerce').fillna(0.0)
    d['적요U'] = d['적요'].fillna('').astype(str).map(_커_글씨)
    d['세부'] = d.get('현금흐름-세부', pd.Series([''] * len(d), index=d.index)).fillna('').astype(str)
    return d.reset_index(drop=True)


def _커_글씨(s):
    """비교용 — 대문자 · 문장부호는 빈칸 · 빈칸은 하나로."""
    return ' ' + re.sub(r'\s+', ' ', re.sub(r'[^0-9A-Z가-힣]+', ' ', str(s).upper())).strip() + ' '


def _커_열쇠말(이름, 별칭표):
    """통장 적요에서 이 거래처를 찾을 글자들 — 별칭표에 있으면 그것, 없으면 거래처 이름 전체."""
    u = 커_이름정리(이름).upper()
    for k, v in 별칭표.items():
        if u.startswith(k.upper()) or k.upper() in u:
            return [_커_글씨(x) for x in v]
    return [_커_글씨(u)] if len(u.strip()) >= 3 else []


def _커_스타일(c):
    return (copy.copy(c.font), copy.copy(c.fill), copy.copy(c.border), copy.copy(c.alignment),
            c.number_format, copy.copy(c.protection))


def _커_입히기(c, s):
    if s is None:
        return
    c.font, c.fill, c.border, c.alignment, c.number_format, c.protection = (copy.copy(x) for x in s)


def _커_옮김(값, 원행, 새행, 칸):
    if isinstance(값, str) and 값.startswith('=') and 원행 != 새행:
        try:
            return Translator(값, origin=f'{칸}{원행}').translate_formula(f'{칸}{새행}')
        except Exception:
            return 값
    return 값


def _커_번호(v):
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else str(f)
    except (TypeError, ValueError):
        return str(v)


def _커_원장시트(W, g, 끝날, 매핑):
    """「26년 원장_raw」 — 계정 머리줄 → Beginning Balance → 거래 → Total for (퀵북 General Ledger List 모양)."""
    # 이전 달 행에서 손으로 넣은 활동분류(Q) · 활동세부(R) · S · T 이어받기
    이전 = {}
    배움 = {}
    for r in range(6, W.max_row + 1):
        b = W.cell(r, 2).value
        if not b or b == 'Beginning Balance' or W.cell(r, 3).value is None:
            continue
        키 = (_커_번호(W.cell(r, 6).value), str(b).strip(), round(float(W.cell(r, 9).value or 0), 2),
              round(float(W.cell(r, 10).value or 0), 2))
        qrst = tuple(W.cell(r, c).value for c in (17, 18, 19, 20))
        이전[키] = qrst
        if qrst[0]:
            배움[(str(b).strip(), 커_이름정리(W.cell(r, 5).value))] = qrst[:2]
    # 줄 모양 (머리 · 기초 · 거래 · 합계) 과 수식 — 지금 시트의 첫 계정 묶음에서 가져옵니다
    머리행 = 6
    tot = next(r for r in range(7, W.max_row + 1) if str(W.cell(r, 1).value or '').startswith('Total for'))
    본 = {'머리': 6, '기초': 7, '거래': 8, '합계': tot}
    모양 = {k: {c: _커_스타일(W.cell(r, c)) for c in range(1, 23)} for k, r in 본.items()}
    수식 = {k: {c: W.cell(r, c).value for c in range(11, 23)
               if isinstance(W.cell(r, c).value, str) and str(W.cell(r, c).value).startswith('=')}
            for k, r in 본.items()}
    for k, r in 본.items():               # 분류(O) 수식 — 이자수익 · 외환 계정도 영업외손익으로
        f = 수식[k].get(15)
        if f:
            for 과목 in ('이자수익', '외환차익', '외환차손', '외환차손익', '외화환산손익'):
                if f'"{과목}"' not in f and f'P{r}="잡이익"' in f:
                    f = f.replace(f'P{r}="잡이익"', f'P{r}="잡이익",P{r}="{과목}"', 1)
            수식[k][15] = f
    높이 = {k: W.row_dimensions[r].height for k, r in 본.items()}
    for rng in list(W.merged_cells.ranges):
        if rng.min_row >= 6:
            W.unmerge_cells(str(rng))
    W.delete_rows(6, W.max_row)

    def 쓰기(r, 종류, 값들):
        for c in range(1, 23):
            _커_입히기(W.cell(r, c), 모양[종류].get(c))
        for c, f in 수식[종류].items():
            W.cell(r, c, _커_옮김(f, 본[종류], r, openpyxl.utils.get_column_letter(c)))
        for c, v in 값들.items():
            W.cell(r, c, v)
        if 높이[종류]:
            W.row_dimensions[r].height = 높이[종류]

    def 코드(a):
        m = re.match(r'^(\d+)', a)
        return m.group(1) if m else ('99999' if a.startswith('Opening') else 'Z' + a)

    def 차변성(a):
        k = 코드(a)
        if k.startswith('1'):
            return True
        if k.startswith('5') and not k.startswith(('511', '531')):
            return True
        return False

    def 활동(a, 이름):
        k = 코드(a)
        if k.startswith('511'):
            return ('상품 판매', '상품매출') if '상품' in a else (('중개 무역', '중개매출') if '중개' in a else ('IOR 서비스', 'IOR매출'))
        if k.startswith('521'):
            return ('상품 원가', '상품매출원가')
        if k.startswith(('53', '54')):
            return ('영업외손익', a)
        if k.startswith('55'):
            return ('법인세비용', a)
        if k.startswith('52'):
            if (a, 이름) in 배움:
                return 배움[(a, 이름)]
            같은이름 = next((v for (aa, nn), v in 배움.items() if nn == 이름 and v[0]), None)
            if 같은이름 and 같은이름[0] == '외주용역비':
                return ('외주용역비', 이름)
            return ('기타', '기타 (1,000불 미만 거래처)')
        return (None, None)

    t = g[g['일자'] <= 끝날]
    r = 6
    새분류 = 0
    for a in sorted(t['계정'].unique(), key=코드):
        x = t[t['계정'] == a].sort_values(['일자', 'TID'], kind='stable')
        쓰기(r, '머리', {1: a})
        r += 1
        쓰기(r, '기초', {2: 'Beginning Balance', 8: 0})
        r += 1
        첫, 잔 = r, 0.0
        IS = 매핑.get(a.lower(), ('', ''))[0] == 'IS'
        for _, y in x.iterrows():
            잔 += (y['차'] - y['대']) if 차변성(a) else (y['대'] - y['차'])
            이름 = 커_이름정리(y['거래처원'])
            값 = {2: a, 3: y['일자'].to_pydatetime(), 4: y['적요'] or None, 5: y['거래처원'] or None,
                  6: y['TID'], 7: (None if pd.isna(y['Num']) else y['Num']) if 'Num' in y else None,
                  8: round(잔, 2), 9: (y['차'] or None), 10: (y['대'] or None)}
            키 = (_커_번호(y['TID']), a, round(float(y['차']), 2), round(float(y['대']), 2))
            if 키 in 이전:
                q, rr, s_, t_ = 이전[키]
                값.update({17: q, 18: rr, 19: s_, 20: t_})
            elif IS:
                q, rr = 활동(a, 이름)
                값.update({17: q, 18: rr})
                새분류 += 1
            쓰기(r, '거래', 값)
            r += 1
        쓰기(r, '합계', {1: f'Total for {a}', 8: round(잔, 2), 9: f'=SUM(I{첫}:I{r - 1})', 10: f'=SUM(J{첫}:J{r - 1})'})
        r += 1
    W.auto_filter.ref = f'A5:V{r - 1}'
    return 새분류, int(len(t))


def _커_채권채무(ws, g, 통장, 끝날, 매핑):
    """채권채무 시트 ① 매출채권 Aging · ③ 채무 표를 원장(+통장)으로 새로 씁니다. ② 미청구 대납비용은 그대로."""
    안내 = []
    원장끝 = g['일자'].max()
    통장끝 = 통장['일자'].max() if 통장 is not None and len(통장) else None
    기준일 = max([x for x in (원장끝, 통장끝) if x is not None])
    다음날 = 끝날 + pd.Timedelta(days=1)
    md = lambda d: f'{d.month}/{d.day}'
    한글 = lambda a: str(매핑.get(a.lower(), ('', ''))[1] or '')

    # ── 지금 시트에서 구역 찾기
    def 찾기(앞):
        return next((r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 2).value or '').startswith(앞)), None)
    s1 = next((r for r in range(1, ws.max_row + 1) if re.match(r'^1[.\-]', str(ws.cell(r, 2).value or '').strip())), None)
    s2 = next((r for r in range(1, ws.max_row + 1) if re.match(r'^2\.', str(ws.cell(r, 2).value or '').strip())
               and any(k in str(ws.cell(r, 2).value) for k in ('미청구', '선급금'))), None)
    s3 = 찾기('3. 채무')
    if not (s1 and s2 and s3):
        return ['채권채무 시트의 구역(1. 매출채권 · 2. 선급금(미청구) · 3. 채무)을 찾지 못해 그대로 두었습니다.']

    def 합계줄(시작):
        return next(r for r in range(시작 + 2, ws.max_row + 1) if str(ws.cell(r, 2).value or '').strip() == '합계')
    t3 = 합계줄(s3)
    d3 = (s3 + 2, t3 - 1)

    # ── 줄마다 값 · 모양 · 병합 · 높이 사진
    병합 = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_row == rng.max_row:
            병합.setdefault(rng.min_row, []).append((rng.min_col, rng.max_col))
    최대열 = max(ws.max_column, 16)
    def 사진(r):
        return {'값': {c: ws.cell(r, c).value for c in range(1, 최대열 + 1)},
                '모양': {c: _커_스타일(ws.cell(r, c)) for c in range(1, 최대열 + 1)},
                '병합': 병합.get(r, []), '높이': ws.row_dimensions[r].height, '원행': r}
    줄 = {r: 사진(r) for r in range(1, ws.max_row + 1)}

    # ── 이전 ① 구역의 비고 이어받기 : (일자, 차변, 대변) → 비고  (사용자가 적은 비고를 다음 달에도 유지)
    def _날(v):
        if isinstance(v, (datetime.datetime, datetime.date)):
            return pd.Timestamp(v).normalize()
        if isinstance(v, (int, float)) and 20000 < v < 80000:
            return pd.Timestamp('1899-12-30') + pd.Timedelta(days=float(v))
        return None
    def _수(v):
        return round(float(v), 2) if isinstance(v, (int, float)) else 0.0
    이전비고, 지금제목 = {}, ''
    for r in range(s1, s2):
        v = 줄[r]['값']
        if isinstance(v.get(2), str) and re.match(r'^\d', v.get(2).strip()):
            지금제목 = v.get(2)
            continue
        d_ = _날(v.get(2))
        if d_ is None or not v.get(10):
            continue
        if '기준일' in 지금제목:
            이전비고[('요약', '선수금' if '선수금' in 지금제목 else '매출채권', 커_이름정리(v.get(3)).upper())] = v.get(10)
        else:
            이전비고[(d_, _수(v.get(6)), _수(v.get(7)))] = v.get(10)
    이전선급 = {}
    for r in range(s2 + 2, s3):
        v = 줄[r]['값']
        if v.get(2) and _날(v.get(4)) is not None and v.get(8):
            이전선급[(커_이름정리(v.get(2)).upper(), _날(v.get(4)), round(_수(v.get(5)), 0))] = v.get(8)
    # 4. 팩토링 건별 원장(③)에서 건 번호 찾기 : (일자, 금액) → 'D1-1 · 브랜드'
    건이름, 건번호 = {}, {}
    for r, x in 줄.items():
        v = x['값']
        if re.fullmatch(r'D\d+', str(v.get(2) or '')) and v.get(3) and not isinstance(v.get(3), (datetime.datetime, datetime.date)):
            건이름[str(v[2])] = str(v[3])
    열매출, 열회수 = 9, 10            # 건별 원장 머리 줄에서 「매출 · 회수」 칸 찾기 (사용자가 칸을 지우면 위치가 바뀜)
    for r, x in 줄.items():
        v = x['값']
        if str(v.get(2) or '').strip() == '건' and str(v.get(4) or '').strip() == '구분':
            for c, val in v.items():
                if str(val or '').strip() == '매출': 열매출 = c
                if str(val or '').strip() == '회수': 열회수 = c
            break
    for r, x in 줄.items():
        v = x['값']
        if v.get(4) in ('매출', '회수') and re.match(r'D\d+', str(v.get(2) or '')) and _날(v.get(3)) is not None:
            금 = v.get(열매출) if v.get(4) == '매출' else v.get(열회수)
            if isinstance(금, (int, float)):
                기본 = re.match(r'D\d+', str(v[2])).group(0)
                건번호[(_날(v.get(3)), round(float(금), 0))] = f'{v[2]} · {건이름.get(기본, "")}'.rstrip(' ·')

    이전3 = []
    for r in range(d3[0], d3[1] + 1):
        v = 줄[r]['값']
        if v.get(2):
            이전3.append((str(v.get(2)), str(v.get(3) or ''), v.get(8)))

    # ── ① 매출채권 · 선수금 : 팩토링 거래처는 원장처럼 건별, 그 외는 거래처별 집계
    ar계정 = [a for a in g['계정'].unique() if 한글(a) == '매출채권'] or [a for a in g['계정'].unique() if a.startswith('11501')]
    선수계정 = [a for a in g['계정'].unique() if 한글(a) == '선수금'] or [a for a in g['계정'].unique() if a.startswith('21511')]
    def _이름열(x):
        return x['고객'].where(x['고객'].notna(), x['거래처원']).fillna('(거래처 표시 없음)').map(커_이름정리)
    def _팩토링(n):
        u = 커_이름정리(n).upper()
        return any(k.upper() in u for k in 커_팩토링거래처)
    def _뒤통장(이름):
        """원장이 끝난 뒤 통장 입금 중 이 거래처 것 (적요 글자로 찾음)."""
        열쇠 = _커_열쇠말(이름, 커_고객별칭)
        if 통장 is None or not 열쇠:
            return 통장.iloc[0:0] if 통장 is not None else None
        u = 통장[(통장['일자'] > 원장끝) & (통장['입금'] > 0)]
        if '채권회수' in set(u['세부']):
            u = u[u['세부'].eq('채권회수')]
        return u[u['적요U'].apply(lambda s: any(k in s for k in 열쇠))]
    전체 = []
    for 종류, 계정들 in (('매출채권', ar계정), ('선수금', 선수계정)):
        x = g[g['계정'].isin(계정들)].copy()
        if not len(x):
            continue
        x['이름'] = _이름열(x)
        x['종류'] = 종류
        전체.append(x)
    전체 = pd.concat(전체) if 전체 else g.iloc[0:0].assign(이름=None, 종류=None)
    def _순서(n):
        u = 커_이름정리(n).upper()
        return next((i for i, k in enumerate(커_팩토링거래처) if k.upper() in u), 99), n
    팩토링이름 = sorted({n for n in 전체['이름'] if _팩토링(n)}, key=_순서)   # 팩토링 거래처 목록에 적은 순서대로
    원장1 = []          # 팩토링 거래처 : [(거래처, 종류, 줄들, 결산후줄들)]
    for 이름 in 팩토링이름:
        for 종류 in ('매출채권', '선수금'):
            x = 전체[(전체['이름'] == 이름) & (전체['종류'] == 종류)].sort_values(['일자', 'TID'], kind='stable')
            전 = x[x['일자'] <= 끝날]
            뒤 = []
            if 종류 == '매출채권':
                for _, y in x[(x['일자'] > 끝날) & (x['대'] > 0)].iterrows():
                    뒤.append((y['일자'], y['적요'], float(y['대']), f'원장 TID {_커_번호(y["TID"])}'))
                u = _뒤통장(이름)
                if u is not None:
                    for _, y in u.iterrows():
                        뒤.append((y['일자'], str(y.get('적요') or ''), float(y['입금']), '통장 · 장부 반영 전'))
            원장1.append((이름, 종류, 전, 뒤))
    집계 = []           # 그 외 거래처
    for (이름, 종류), x in 전체[~전체['이름'].isin(팩토링이름)].groupby(['이름', '종류'], sort=False):
        전 = x[x['일자'] <= 끝날]
        차, 대 = float(전['차'].sum()), float(전['대'].sum())
        잔 = 차 - 대 if 종류 == '매출채권' else 대 - 차
        if abs(차) < 0.005 and abs(대) < 0.005:
            continue
        뒤 = 0.0
        if 종류 == '매출채권':
            뒤 = float(x[x['일자'] > 끝날]['대'].sum())
            u = _뒤통장(이름)
            if u is not None:
                뒤 += float(u['입금'].sum())
        회수후 = min(max(잔, 0.0), 뒤)
        비 = []
        if 종류 == '매출채권' and 뒤 - 회수후 > 0.5:
            비.append(f'결산 후 입금 {뒤:,.0f} 중 {뒤 - 회수후:,.0f} 은 결산 뒤 청구분 회수로 보임 — 확인 필요')
        if 종류 == '매출채권' and 잔 < -0.5:
            비.append('청구보다 더 받음 — 다음 청구와 상계 또는 선수금 대체 필요')
        # 적요에 팩토링 거래처 이름이 있는데 퀵북 거래처가 다른 줄 (확인용)
        for _, y in 전.iterrows():
            for k in 커_팩토링거래처:
                if k.upper() in str(y['적요']).upper():
                    비.append(f'{y["일자"]:%m/%d} TID {_커_번호(y["TID"])} 「{y["적요"]}」 {max(y["차"], y["대"]):,.0f} 은 '
                              f'적요는 {k} 인데 퀵북 거래처가 {이름} — 확인 필요')
        집계.append({'거래처': 이름, '종류': 종류, '차': 차, '대': 대, '잔': 잔, '뒤': 회수후, '비고': ' · '.join(비)})
    집계.sort(key=lambda d: (0 if d['종류'] == '매출채권' else 1, -d['잔'], d['거래처']))
    if 원장끝 > 끝날 and float(전체[(전체['종류'] == '매출채권') & (전체['일자'] > 끝날)]['차'].sum()) > 0.5:
        안내.append(f'원장에 결산월 뒤 청구 {float(전체[(전체["종류"] == "매출채권") & (전체["일자"] > 끝날)]["차"].sum()):,.0f} 이 있어 채권 표에는 넣지 않았습니다.')

    # ── ③ 채무 : 미지급금 · 단기차입금 거래처별 결산월 말 잔액
    채무계정 = [a for a in g['계정'].unique() if 한글(a) in ('미지급금', '단기차입금')]
    행3, 건너뜀 = [], []
    for a in sorted(채무계정):
        x = g[g['계정'] == a]
        이름열 = x['공급'].fillna(x['거래처원']).fillna('(거래처 표시 없음)').map(커_이름정리)
        for _k, y in x.groupby(이름열.str.upper()):
            이름 = 이름열[y.index[0]]
            전 = y[y['일자'] <= 끝날]
            잔 = float(전['대'].sum() - 전['차'].sum())
            if abs(잔) < 0.5:
                continue
            if 한글(a) == '미지급금' and 잔 < 0:      # 미지급금 차변잔액은 결산조정분개로 자산(선급비용 · 미수금)으로 감
                건너뜀.append(f'{커_이름정리(이름열[y.index[0]])} {잔:,.0f}')
                continue
            뒤원장 = float(y[y['일자'] > 끝날]['차'].sum())
            뒤통장, 열쇠 = 0.0, _커_열쇠말(이름, 커_거래처별칭)
            if 통장 is not None and 열쇠:
                u = 통장[(통장['일자'] > 원장끝) & (통장['출금'] > 0)]
                뒤통장 = float(u[u['적요U'].apply(lambda s: any(k in s for k in 열쇠))]['출금'].sum())
            지급 = min(max(잔, 0.0), 뒤원장 + 뒤통장)
            행3.append({'거래처': 이름, '계정': 한글(a), '잔액': 잔, '지급': 지급})
    # 이전 줄의 이름 · 비고 이어받기 (낱말이 가장 많이 겹치는 줄)
    def 낱말(s):
        return set(re.findall(r'[0-9A-Za-z가-힣]+', str(s).upper()))
    for d in 행3:
        best, 점 = None, 0.0
        for n, 계, 비 in 이전3:
            if 계 != d['계정']:
                continue
            a_, b_ = 낱말(n), 낱말(d['거래처'])
            j = len(a_ & b_) / max(len(a_ | b_), 1)
            if j > 점:
                best, 점 = (n, 비), j
        if best and 점 >= 0.3:
            d['거래처'], d['비고'] = best[0], best[1]
        else:
            d['비고'] = None
    순서 = {'미지급금': 0, '단기차입금': 1}
    행3.sort(key=lambda d: (순서.get(d['계정'], 2), -d['잔액']))

    새모양 = any(re.match(r'^1-\d+\.', str(줄[r]['값'].get(13) or '').strip()) for r in range(s1, s2))
    if 새모양:
        # ══ 사용자가 다듬은 모양(왼쪽 집계 · 선급금 · 채무 / 오른쪽 팩토링 세부 원장)을 칸 · 줄 모양 그대로 다시 씀 (2026-11-02)
        최대 = max(줄)
        def _글(r, c):
            return str(줄[r]['값'].get(c) or '').strip() if r in 줄 else ''
        def _표(r제목, 열):
            """제목 줄에서 표 범위 : 머리 · 본문 줄들 · 합계 · 메모."""
            t = {'제목': r제목, '머리': r제목 + 1, '본문': [], '합계': None, '메모': None, '글': _글(r제목, 열)}
            r = r제목 + 2
            while r <= 최대 and 줄[r]['값'].get(열) is not None and _글(r, 열) != '합계' \
                    and not _글(r, 열).startswith('※') and not (isinstance(줄[r]['값'].get(열), str) and re.match(r'^\d(-\d+)*\.', _글(r, 열))):
                t['본문'].append(r)
                r += 1
            if r <= 최대 and _글(r, 열) == '합계':
                t['합계'] = r
                r += 1
            if r <= 최대 and _글(r, 열).startswith('※'):
                t['메모'] = r
                r += 1
            t['끝'] = r - 1
            return t
        왼, 오 = [], []
        for r in range(s1, 최대 + 1):
            if re.match(r'^\d(-\d+)*\.', _글(r, 2)) and _글(r + 1, 2) in ('일자', '지급일', '거래처'):
                왼.append(_표(r, 2))
            if re.match(r'^\d(-\d+)*\.', _글(r, 13)) and _글(r + 1, 13) == '일자':
                오.append(_표(r, 13))
        def _종류(t):
            g_ = t['글']
            if re.match(r'^1\.', g_): return '요약AR'
            if re.match(r'^1-1\.', g_): return '요약선수'
            if re.match(r'^1-\d+-1\.', g_): return '세부선수'
            if re.match(r'^2\.', g_): return '선급'
            if re.match(r'^3\.', g_): return '채무'
            return '기타'
        틀 = {}
        for t in 왼:
            틀.setdefault(_종류(t), t)
        왼끝 = max([t['끝'] for t in 왼] + [s1])
        꼬리시작 = 왼끝 + 1                      # 그 아래(건별 원장 등)는 그대로 옮김
        def _간격(t, 다음):
            return max(다음['제목'] - t['끝'] - 1, 0) if 다음 else 2
        간 = {}
        for i, t in enumerate(왼):
            간.setdefault(_종류(t), _간격(t, 왼[i + 1] if i + 1 < len(왼) else None))
        오간 = _간격(오[0], 오[1]) if len(오) > 1 else 1
        날짜글 = lambda s: re.sub(r'기준일 : \d{4}\.\d{2}\.\d{2}', f'기준일 : {끝날:%Y.%m.%d}',
                             re.sub(r'기간 : \d{4}\.\d{2}\.\d{2} ~ \d{4}\.\d{2}\.\d{2}', f'기간 : {끝날.year}.01.01 ~ {끝날:%Y.%m.%d}', s))

        # 선급금 열린 금액 (원장 선급금 계정 · 지급처별로 먼저 지급한 것부터 청구 · 대체 금액을 지움)
        선급계정 = [a for a in g['계정'].unique() if 한글(a) == '선급금'] or [a for a in g['계정'].unique() if a.startswith('11801')]
        sx = g[g['계정'].isin(선급계정) & (g['일자'] <= 끝날)].copy()
        sx['이름'] = sx['거래처원'].fillna('(거래처 표시 없음)').map(커_이름정리)
        sx = sx.sort_values(['일자', 'TID'], kind='stable')
        열린, 남은대 = [], []
        for 이름, x in sx.groupby('이름', sort=False):
            큐 = [{'이름': 이름, '일자': y['일자'], '적요': y['적요'], '지급': float(y['차']), '청구': 0.0}
                  for _, y in x[x['차'] > 0].iterrows()]
            돈 = float(x['대'].sum())
            for q in 큐:
                뺄 = min(돈, q['지급']); q['청구'] += 뺄; 돈 -= 뺄
            if 돈 > 0.005:
                끝줄 = x[x['대'] > 0].iloc[-1]
                남은대.append({'이름': 이름, '일자': 끝줄['일자'], '적요': 끝줄['적요'], '금액': 돈})
            열린 += 큐
        for z in 남은대:
            짝 = sx[(sx['차'] - z['금액']).abs() < 0.01]
            if len(짝):
                for q in [q for q in 열린 if q['이름'] == 짝['이름'].iloc[0]]:
                    뺄 = min(z['금액'], q['지급'] - q['청구']); q['청구'] += 뺄; z['금액'] -= 뺄
                    if z['금액'] <= 0.005:
                        break
        선급행 = [q for q in 열린 if q['지급'] - q['청구'] > 0.005] + \
                 [{'이름': z['이름'], '일자': z['일자'], '적요': z['적요'], '지급': 0.0, '청구': z['금액'], '음': True}
                  for z in 남은대 if z['금액'] > 0.005]
        선급잔 = float(sx['차'].sum() - sx['대'].sum())
        # ① 집계 자료
        요약 = []
        for (이름, 종류), x in 전체.groupby(['이름', '종류'], sort=False):
            전 = x[x['일자'] <= 끝날]
            차, 대 = float(전['차'].sum()), float(전['대'].sum())
            if abs(차) >= 0.005 or abs(대) >= 0.005:
                요약.append({'거래처': 이름, '종류': 종류, '차': 차, '대': 대, '팩': 이름 in 팩토링이름})
        요약.sort(key=lambda d: (0 if d['팩'] else 1, _순서(d['거래처']) if d['팩'] else (0, ''),
                                -((d['차'] - d['대']) if d['종류'] == '매출채권' else (d['대'] - d['차'])), d['거래처']))
        세부번호 = {n: f'1-{i + 2}' for i, n in enumerate(팩토링이름)}

        # ── 시트 비우고 다시 쓰기
        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))
        ws.delete_rows(1, ws.max_row + 5)
        def 복사(원, 새r, c1, c2, 값=None, 높이=True):
            s_ = 줄[원]
            for c in range(c1, c2 + 1):
                if c in s_['모양']:
                    _커_입히기(ws.cell(새r, c), s_['모양'][c])
                if 값 is not None and c in 값:
                    ws.cell(새r, c).value = 값[c]
                else:
                    ws.cell(새r, c).value = _커_옮김(s_['값'].get(c), 원, 새r, openpyxl.utils.get_column_letter(c))
            for a, b in s_['병합']:
                if a >= c1 and b <= c2:
                    ws.merge_cells(start_row=새r, start_column=a, end_row=새r, end_column=b)
            if 높이 and s_['높이']:
                ws.row_dimensions[새r].height = s_['높이']
        for r in range(1, s1):                       # 머리 (제목 · 팩토링 표시 · 오른쪽 위 메모)
            복사(r, r, 1, 최대열)
        L = lambda c: openpyxl.utils.get_column_letter(c)
        def 표쓰기(t, 시작, 행들, 열, 범위, 키함수, 값함수, 합계식=None, 제목글=None, 메모글=None, 높이=True):
            """t 견본 표를 시작 줄부터 다시 씀 — 같은 키의 이전 줄이 있으면 그 줄 모양 · 글을 그대로 씀."""
            c1, c2 = 범위
            r = 시작
            복사(t['제목'], r, c1, c2, {열: 제목글 if 제목글 else 날짜글(t['글'])}, 높이); r += 1
            복사(t['머리'], r, c1, c2, None, 높이); r += 1
            이전 = {}
            for rr in t['본문']:
                k = 키함수(줄[rr]['값'], True)
                if k is not None:
                    이전.setdefault(k, rr)
            첫 = r
            본1 = t['본문'][0] if t['본문'] else t['머리']
            본끝 = t['본문'][-1] if t['본문'] else 본1
            for i, d in enumerate(행들):
                k = 키함수(d, False)
                원 = 이전.get(k)
                마지막 = i == len(행들) - 1
                견 = 원 or (본끝 if (마지막 and t['합계'] is None) else 본1)
                v = 값함수(d, r, r == 첫, 줄[원]['값'] if 원 else None)
                복사(견, r, c1, c2, v, 높이)
                if 마지막 and t['합계'] is None and 원 and 원 != 본끝:      # 표 끝줄 아래 선은 견본 끝줄 것
                    for c in range(c1, c2 + 1):
                        b_ = copy.copy(ws.cell(r, c).border); e_ = 줄[본끝]['모양'].get(c)
                        if e_:
                            ws.cell(r, c).border = openpyxl.styles.Border(left=b_.left, right=b_.right, top=b_.top, bottom=copy.copy(e_[2].bottom))
                r += 1
            끝 = r - 1
            if t['합계'] is not None:
                복사(t['합계'], r, c1, c2, {c: f'=SUM({L(c)}{첫}:{L(c)}{끝})' for c in (합계식 or [])}, 높이); r += 1
            if t['메모'] is not None:
                복사(t['메모'], r, c1, c2, {열: 메모글} if 메모글 else None, 높이); r += 1
            return r - 1
        def _k날(v):
            return _날(v) if not isinstance(v, pd.Timestamp) else v.normalize()

        # 왼쪽
        r = s1
        def 다음(r, 종류):
            return r + 1 + 간.get(종류, 2)
        if '요약AR' in 틀:
            행 = [d for d in 요약 if d['종류'] == '매출채권']
            끝r = 표쓰기(틀['요약AR'], r, 행, 2, (1, 11),
                        lambda v, 옛: (커_이름정리(v.get(3)).upper() if 옛 else v['거래처'].upper()),
                        lambda d, rr, 첫, 옛: {2: pd.Timestamp(끝날).to_pydatetime(), 3: d['거래처'], 4: f'1~{끝날.month}월 누적',
                                              6: round(d['차'], 2) or None, 7: round(d['대'], 2) or None, 8: f'=F{rr}-G{rr}',
                                              9: (옛.get(9) if 옛 and 옛.get(9) else (f'팩토링 — 세부 {세부번호[d["거래처"]]} 참조' if d['팩'] else None))},
                        합계식=(6, 7, 8))
            r = 다음(끝r, '요약AR')
        행 = [d for d in 요약 if d['종류'] == '선수금']
        if '요약선수' in 틀 and 행:
            끝r = 표쓰기(틀['요약선수'], r, 행, 2, (1, 11),
                        lambda v, 옛: (커_이름정리(v.get(3)).upper() if 옛 else v['거래처'].upper()),
                        lambda d, rr, 첫, 옛: {2: pd.Timestamp(끝날).to_pydatetime(), 3: d['거래처'], 4: f'1~{끝날.month}월 누적',
                                              6: round(d['차'], 2) or None, 7: round(d['대'], 2) or None, 8: f'=G{rr}-F{rr}',
                                              9: (옛.get(9) if 옛 and 옛.get(9) else (f'팩토링 — 세부 {세부번호[d["거래처"]]} 참조' if d['팩'] else None))},
                        합계식=(6, 7, 8))
            r = 다음(끝r, '요약선수')
        def 원장값(n, 수익쪽, 열들):
            """원장 줄 값 — 열들 = (일자, 거래처, 적요, 차, 대, 잔, 비고) 칸 번호."""
            cD, cN, cJ, cDr, cCr, cB, cR = 열들
            def f(d, rr, 첫, 옛):
                if d.get('뒤'):
                    식 = f'={L(cB)}{rr - 1}-{L(cCr)}{rr}' if not 첫 else f'=-{L(cCr)}{rr}'
                    return {cD: pd.Timestamp(d['일자']).to_pydatetime(), cN: n, cJ: (옛.get(cJ) if 옛 else f'(결산 후) {d["적요"]}'),
                            cDr: None, cCr: round(d['대'], 2), cB: 식, cR: (옛.get(cR) if 옛 and 옛.get(cR) else d['비고'])}
                if 수익쪽:
                    식 = f'={L(cCr)}{rr}-{L(cDr)}{rr}' if 첫 else f'={L(cB)}{rr - 1}+{L(cCr)}{rr}-{L(cDr)}{rr}'
                else:
                    식 = f'={L(cDr)}{rr}-{L(cCr)}{rr}' if 첫 else f'={L(cB)}{rr - 1}+{L(cDr)}{rr}-{L(cCr)}{rr}'
                return {cD: pd.Timestamp(d['일자']).to_pydatetime(), cN: n, cJ: (옛.get(cJ) if 옛 and 옛.get(cJ) else (d['적요'] or None)),
                        cDr: round(d['차'], 2) or None, cCr: round(d['대'], 2) or None, cB: 식,
                        cR: (옛.get(cR) if 옛 and 옛.get(cR) else d['비고'])}
            return f
        def 원장행(n, 전, 뒤):
            행 = []
            for _, y in 전.iterrows():
                키 = (pd.Timestamp(y['일자']).normalize(), round(float(y['차']), 2), round(float(y['대']), 2))
                건 = 건번호.get((키[0], round(max(float(y['차']), float(y['대'])), 0)))
                비 = ' · '.join(x for x in (건, f'TID {_커_번호(y["TID"])}') if x)
                for k in 커_팩토링거래처:
                    if k.upper() in str(y['적요']).upper() and k.upper() not in n.upper():
                        비 += f' · 적요는 {k} 인데 퀵북 거래처가 {n} — 확인 필요'
                행.append({'일자': y['일자'], '적요': y['적요'], '차': float(y['차']), '대': float(y['대']), '비고': 비})
            for (d_, 적, 금, 근) in sorted(뒤, key=lambda t: t[0]):
                행.append({'일자': d_, '적요': 적, '차': 0.0, '대': 금, '비고': 근, '뒤': True})
            return 행
        def 원장키(열들):
            cD, cN, cJ, cDr, cCr, cB, cR = 열들
            def f(v, 옛):
                if 옛:
                    d_ = _날(v.get(cD))
                    return None if d_ is None else (d_, round(_수(v.get(cDr)), 2), round(_수(v.get(cCr)), 2))
                return (pd.Timestamp(v['일자']).normalize(), round(v['차'], 2), round(v['대'], 2))
            return f
        왼열 = (2, 3, 4, 6, 7, 8, 9)
        if '세부선수' in 틀 or '요약선수' in 틀:
            견선 = 틀.get('세부선수') or 틀.get('요약선수')
            for n in 팩토링이름:
                for (n2, 종류, 전, 뒤) in 원장1:
                    if n2 != n or 종류 != '선수금' or not len(전):
                        continue
                    끝r = 표쓰기(견선, r, 원장행(n, 전, []), 2, (1, 11), 원장키(왼열), 원장값(n, True, 왼열),
                                제목글=f'{세부번호[n]}-1. {n} 선수금 (기간 : {끝날.year}.01.01 ~ {끝날:%Y.%m.%d})')
                    r = 다음(끝r, '세부선수')
        if '선급' in 틀:
            선열 = (2, 3, 4, 6, 7, 8, 9)       # 지급일 · 지급처 · 내용 · 지급액 · 청구액 · 미청구 · 비고
            def 선키(v, 옛):
                if 옛:
                    d_ = _날(v.get(2))
                    return None if d_ is None else (커_이름정리(v.get(3)).upper(), d_, round(_수(v.get(6)), 0))
                return (v['이름'].upper(), pd.Timestamp(v['일자']).normalize(), round(v['지급'], 0))
            def 선값(q, rr, 첫, 옛):
                비 = 옛.get(9) if 옛 and 옛.get(9) else ('지급 없이 대체된 금액 — 확인 필요' if q.get('음') else None)
                return {2: pd.Timestamp(q['일자']).to_pydatetime(), 3: q['이름'], 4: q['적요'] or None,
                        6: round(q['지급'], 2), 7: round(q['청구'], 2), 8: f'=F{rr}-G{rr}', 9: 비}
            끝r = 표쓰기(틀['선급'], r, 선급행, 2, (1, 11), 선키, 선값, 합계식=(6, 7, 8),
                        메모글=f'※ 합계 {선급잔:,.0f} = BS 선급금 ({md(끝날)} 원장) · 지급처별로 먼저 지급한 것부터 청구 · 대체 금액을 지움')
            r = 다음(끝r, '선급')
        if '채무' in 틀:
            def 채키(v, 옛):
                if 옛:
                    return (str(v.get(2) or '').upper(), str(v.get(3) or ''))
                return (str(v['거래처']).upper(), v['계정'])
            def 채값(d, rr, 첫, 옛):
                return {2: d['거래처'], 3: d['계정'], 5: round(d['잔액'], 2), 6: round(d['지급'], 2), 7: f'=E{rr}-F{rr}',
                        8: (옛.get(8) if 옛 and 옛.get(8) else d.get('비고'))}
            끝r = 표쓰기(틀['채무'], r, 행3, 2, (1, 11), 채키, 채값, 합계식=(5, 6, 7),
                        메모글=(f'※ {md(다음날)} 이후 지급 = 원장({md(원장끝)}까지) 차변'
                               + (f' + 통장({md(통장끝)}까지) 출금' if 통장끝 is not None else '') + ' · 결산월 말 잔액을 넘는 지급은 넣지 않았습니다.'))
            r = 끝r
        왼새끝 = r
        # 오른쪽 — 팩토링 매출채권 세부 원장
        r = s1
        오열 = (13, 14, 15, 17, 18, 19, 20)
        for n in 팩토링이름:
            견 = next((t for t in 오 if n.upper() in t['글'].upper()), 오[0])
            for (n2, 종류, 전, 뒤) in 원장1:
                if n2 != n or 종류 != '매출채권' or (not len(전) and not 뒤):
                    continue
                끝r = 표쓰기(견, r, 원장행(n, 전, 뒤), 13, (12, 최대열), 원장키(오열), 원장값(n, False, 오열),
                            제목글=f'{세부번호[n]}. {n} 매출채권 (기간 : {끝날.year}.01.01 ~ {끝날:%Y.%m.%d})', 높이=False)
                r = 끝r + 1 + 오간
        오새끝 = r
        # 꼬리 (건별 원장 등) — 왼쪽 · 오른쪽 중 더 긴 쪽 아래로, 원래 간격 그대로
        꼬리새 = max(왼새끝 + (꼬리시작 - 왼끝), 오새끝 + 1)
        for rr in range(꼬리시작, 최대 + 1):
            복사(rr, 꼬리새 + (rr - 꼬리시작), 1, 최대열)
        b3 = ws.max_row + 10
    else:
        # ── 시트 다시 쓰기
        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))
        ws.delete_rows(1, ws.max_row + 5)

        새 = [1]
        def 놓기(s, 값=None, 원행=None):
            r = 새[0]
            for c, st in s['모양'].items():
                _커_입히기(ws.cell(r, c), st)
            원 = 원행 or s['원행']
            for c, v in s['값'].items():
                if 값 is not None and c in 값:
                    continue
                ws.cell(r, c, _커_옮김(v, 원, r, openpyxl.utils.get_column_letter(c)))
            for c, v in (값 or {}).items():
                ws.cell(r, c, v)
            for a, b in s['병합']:
                ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
            if s['높이']:
                ws.row_dimensions[r].height = s['높이']
            새[0] += 1
            return r

        def 범위고침(r, 원a, 원b, 새a, 새b):
            for c in range(1, 최대열 + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.startswith('='):
                    ws.cell(r, c).value = re.sub(rf'(\$?[A-Z]+\$?){원a}:(\$?[A-Z]+\$?){원b}\b',
                                                 rf'\g<1>{새a}:\g<2>{새b}', v)

        # 모양 견본 — ② 표(달마다 그대로인 표)의 머리 · 본문 · 합계 줄에서 가져옴
        t2 = next(r for r in range(s2 + 2, max(줄) + 1) if str(줄[r]['값'].get(2) or '').strip() == '합계')
        견머리, 견본, 견합 = 줄[s2 + 1]['모양'], 줄[s2 + 2]['모양'], 줄[t2]['모양']
        # 사용자가 다듬어 둔 ① 원장 표(머리 「일자」 줄 · 본문 · 끝줄)가 있으면 그 줄 모양을 칸마다 그대로 씀
        원머리 = next((r for r in range(s1, s2) if str(줄[r]['값'].get(2) or '').strip() == '일자'), None)
        칸별 = False
        if 원머리 and 원머리 + 1 < s2 and 줄[원머리 + 1]['값'].get(2) is not None:
            원끝 = 원머리 + 1
            while 원끝 + 1 < s2 and 줄[원끝 + 1]['값'].get(2) is not None and str(줄[원끝 + 1]['값'].get(2)).strip() != '일자':
                원끝 += 1
            견머리, 견본 = 줄[원머리]['모양'], 줄[원머리 + 1]['모양']
            견끝 = 줄[원끝]['모양']
            견합 = 견끝
            칸별 = True
        빨강 = openpyxl.styles.Font(name='맑은 고딕', size=9, bold=True, color='FFC00000')
        옅음 = openpyxl.styles.PatternFill('solid', fgColor='FFF2F2F2')

        # 1~머리줄
        for r in range(1, s1):
            놓기(줄[r])

        def 칸(r, c, v=None, 모양=None, 형식=None, 굵게=False, 글=None):
            m = 모양 or 견본
            if 칸별:
                _커_입히기(ws.cell(r, c), m.get(c))                        # 사용자 원장 표의 같은 칸 모양
            else:
                _커_입히기(ws.cell(r, c), m.get(5) if 형식 else m.get(2))     # 모든 칸 같은 모양 (글자칸 = B열 견본, 숫자칸 = E열 견본)
            if v is not None:
                ws.cell(r, c).value = v
            if 형식:
                ws.cell(r, c).number_format = 형식
            if 굵게 or 글:
                ws.cell(r, c).font = 글 or openpyxl.styles.Font(name='맑은 고딕', size=9, bold=True)
        숫형 = '#,##0;\\(#,##0\\);"-"'
        def 줄쓰기(값들, 모양, 합칠=(), 숫자칸=(), 왼칸=(), 굵게=False):
            r = 새[0]
            for c in range(2, 16):
                칸(r, c, 값들.get(c), 모양, 숫형 if c in 숫자칸 else None, 굵게)
                ws.cell(r, c).alignment = openpyxl.styles.Alignment(
                    horizontal='right' if c in 숫자칸 else ('left' if c in 왼칸 else 'center'), vertical='center')
            for a, b in 합칠:
                ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
            새[0] += 1
            return r
        def 제목(글, 굵게=True):
            r = 새[0]
            ws.cell(r, 2).value = 글
            ws.cell(r, 2).font = openpyxl.styles.Font(name='맑은 고딕', size=9, bold=굵게)
            새[0] += 1
            return r

        def 표끝(r):
            for c in range(2, 16):          # 표 끝줄 — 사용자 끝줄 모양의 테두리 (없으면 아래 가는 줄)
                b_ = copy.copy(ws.cell(r, c).border)
                아래 = (견끝.get(c)[2].bottom if 칸별 and 견끝.get(c) else None) or openpyxl.styles.Side(style='thin')
                ws.cell(r, c).border = openpyxl.styles.Border(left=b_.left, right=b_.right, top=b_.top, bottom=copy.copy(아래))
        머리1 = {2: '일자', 3: '거래처', 4: '적요', 6: '차변', 7: '대변', 9: '잔액', 10: '비고'}
        합1 = ((4, 5), (7, 8), (10, 15))
        기간 = f'기간 : {끝날.year}.01.01 ~ {끝날:%Y.%m.%d}'
        def 제목합(글, 이름, 식):          # 제목 옆 합계 칸은 사용자가 지워서 제목만 씀 (2026-11-02)
            return 제목(글)
        # ① 집계 — 모든 거래처 (팩토링 거래처 먼저) · 일자 = 결산월 말일
        요약 = []
        for (이름, 종류), x in 전체.groupby(['이름', '종류'], sort=False):
            전 = x[x['일자'] <= 끝날]
            차, 대 = float(전['차'].sum()), float(전['대'].sum())
            if abs(차) < 0.005 and abs(대) < 0.005:
                continue
            요약.append({'거래처': 이름, '종류': 종류, '차': 차, '대': 대, '잔': (차 - 대) if 종류 == '매출채권' else (대 - 차),
                         '팩': 이름 in 팩토링이름})
        요약.sort(key=lambda d: (0 if d['팩'] else 1, _순서(d['거래처']) if d['팩'] else (0, ''), -d['잔'], d['거래처']))
        세부번호 = {}
        번 = 2
        for 이름 in 팩토링이름:
            세부번호[이름] = f'1-{번}'
            번 += 1
        for 종류, 번호, 합이름 in (('매출채권', '1.', '미수채권'), ('선수금', '1-1.', '선수금')):
            행 = [d for d in 요약 if d['종류'] == 종류]
            if 종류 == '선수금' and not 행:
                continue
            r제목 = 제목합(f'{번호} {종류} (기준일 : {끝날:%Y.%m.%d})', 합이름, None)
            줄쓰기(머리1, 견머리, 합1)
            첫 = 새[0]
            for d in 행:
                r = 새[0]
                비 = 이전비고.get(('요약', 종류, d['거래처'].upper()))
                if not 비 and d['팩']:
                    비 = f'팩토링 — 세부 {세부번호[d["거래처"]]} 참조'
                줄쓰기({2: pd.Timestamp(끝날).to_pydatetime(), 3: d['거래처'], 4: f'1~{끝날.month}월 누적',
                       6: round(d['차'], 2) or None, 7: round(d['대'], 2) or None,
                       9: (f'=F{r}-G{r}' if 종류 == '매출채권' else f'=G{r}-F{r}'), 10: 비},
                      견본, 합1, (6, 7, 9), (3, 4, 10))
                ws.cell(r, 2).number_format = 'yyyy-mm-dd'
            끝 = 새[0] - 1
            r = 줄쓰기({2: '합계', 6: f'=SUM(F{첫}:F{끝})', 7: f'=SUM(G{첫}:G{끝})', 9: f'=SUM(I{첫}:I{끝})'}, 견합, ((2, 5), (7, 8), (10, 15)), (6, 7, 9), 굵게=True)
            표끝(r)
            새[0] += 1
        제목(f'※ 팩토링 거래처({", ".join(팩토링이름) or "없음"})는 아래에 건별 원장으로 세부 내역을 보여 줍니다 · '
             f'세부 원장 아래 줄(결산 후)은 통장에 들어온 회수로, 장부에는 아직 반영 전입니다', 굵게=False)
        # ① 세부 — 팩토링 거래처 원장 (원장처럼 발생 · 회수)
        for 이름 in 팩토링이름:
            for (n, 종류, 전, 뒤) in [t for t in 원장1 if t[0] == 이름]:
                if not len(전) and not 뒤:
                    continue
                새[0] += 1
                번호 = 세부번호[이름] + ('.' if 종류 == '매출채권' else '-1.')
                r제목 = 제목합(f'{번호} {n} {종류} ({기간})', '미수채권' if 종류 == '매출채권' else '선수금', None)
                줄쓰기(머리1, 견머리, 합1)
                첫 = 새[0]
                for _, y in 전.iterrows():
                    키 = (pd.Timestamp(y['일자']).normalize(), round(float(y['차']), 2), round(float(y['대']), 2))
                    비 = 이전비고.get(키)
                    if not 비:
                        건 = 건번호.get((키[0], round(max(float(y['차']), float(y['대'])), 0)))
                        비 = ' · '.join(x for x in (건, f'TID {_커_번호(y["TID"])}') if x)
                        for k in 커_팩토링거래처:       # 적요는 다른 팩토링 거래처인데 퀵북 거래처가 이 거래처인 줄
                            if k.upper() in str(y['적요']).upper() and k.upper() not in n.upper():
                                비 += f' · 적요는 {k} 인데 퀵북 거래처가 {n} — 확인 필요'
                    r = 새[0]
                    잔식 = (f'=F{r}-G{r}' if r == 첫 else f'=I{r - 1}+F{r}-G{r}') if 종류 == '매출채권' else \
                           (f'=G{r}-F{r}' if r == 첫 else f'=I{r - 1}+G{r}-F{r}')
                    줄쓰기({2: pd.Timestamp(y['일자']).to_pydatetime(), 3: n, 4: y['적요'] or None,
                           6: round(float(y['차']), 2) or None, 7: round(float(y['대']), 2) or None, 9: 잔식, 10: 비},
                          견본, 합1, (6, 7, 9), (3, 4, 10))
                    ws.cell(r, 2).number_format = 'yyyy-mm-dd'
                for (d_, 적, 금, 근) in sorted(뒤, key=lambda t: t[0]):
                    r = 새[0]
                    키 = (pd.Timestamp(d_).normalize(), 0.0, round(금, 2))
                    비 = 이전비고.get(키) or 근
                    줄쓰기({2: pd.Timestamp(d_).to_pydatetime(), 3: n, 4: f'(결산 후) {적}', 7: round(금, 2),
                           9: (f'=I{r - 1}-G{r}' if r > 첫 else f'=-G{r}'), 10: 비}, 견본, 합1, (6, 7, 9), (3, 4, 10))
                    ws.cell(r, 2).number_format = 'yyyy-mm-dd'
                    ws.cell(r, 10).font = 빨강
                표끝(새[0] - 1)
        새[0] += 2

        # ② 선급금 — 원장 선급금 계정 결산월 말 잔액 (지급처별로 먼저 지급한 것부터 청구 · 대체로 지움)
        선급계정 = [a for a in g['계정'].unique() if 한글(a) == '선급금'] or [a for a in g['계정'].unique() if a.startswith('11801')]
        sx = g[g['계정'].isin(선급계정) & (g['일자'] <= 끝날)].copy()
        sx['이름'] = sx['거래처원'].fillna('(거래처 표시 없음)').map(커_이름정리)
        sx = sx.sort_values(['일자', 'TID'], kind='stable')
        열린, 남은대 = [], []
        for 이름, x in sx.groupby('이름', sort=False):
            큐 = [{'이름': 이름, '일자': y['일자'], '적요': y['적요'], '지급': float(y['차']), '청구': 0.0}
                  for _, y in x[x['차'] > 0].iterrows()]
            돈 = float(x['대'].sum())          # 청구 · 대체(대변) 합계를 먼저 지급한 것부터 지움 (날짜 순서와 무관)
            for q in 큐:
                뺄 = min(돈, q['지급'])
                q['청구'] += 뺄
                돈 -= 뺄
            if 돈 > 0.005:
                끝줄 = x[x['대'] > 0].iloc[-1]
                남은대.append({'이름': 이름, '일자': 끝줄['일자'], '적요': 끝줄['적요'], '금액': 돈})
            열린 += 큐
        for z in 남은대:          # 지급처 이름이 다른 대체(예 : 청구 거래처 이름으로 대체) — 같은 금액을 지급한 지급처의 열린 지급에서 지움
            짝 = sx[(sx['차'] - z['금액']).abs() < 0.01]
            if len(짝):
                대상 = 짝['이름'].iloc[0]
                for q in [q for q in 열린 if q['이름'] == 대상]:
                    뺄 = min(z['금액'], q['지급'] - q['청구'])
                    q['청구'] += 뺄
                    z['금액'] -= 뺄
                    if z['금액'] <= 0.005:
                        break
        선급행 = [q for q in 열린 if q['지급'] - q['청구'] > 0.005] + \
                 [{'이름': z['이름'], '일자': z['일자'], '적요': z['적요'], '지급': 0.0, '청구': z['금액'], '음': True}
                  for z in 남은대 if z['금액'] > 0.005]
        선급잔 = float(sx['차'].sum() - sx['대'].sum())
        머2원 = 줄[s2]['값']
        놓기(줄[s2], {2: f'2. 선급금 (기준일 : {끝날:%Y.%m.%d}) — 지급했으나 거래처에 아직 청구 · 대체하지 않은 금액 (미청구 대납비용 등)'})
        놓기(줄[s2 + 1])
        틀2 = 줄[s2 + 2]
        a2 = 새[0]
        for q in (선급행 or [None]):
            if q is None:
                놓기(틀2, {2: None, 3: '해당 없음', 4: None, 5: 0, 6: 0, 8: None})
                continue
            비 = 이전선급.get((q['이름'].upper(), pd.Timestamp(q['일자']).normalize(), round(q['지급'], 0)))
            if q.get('음'):
                비 = 비 or '지급 없이 대체된 금액 — 확인 필요'
            r = 놓기(틀2, {2: q['이름'], 3: q['적요'] or None, 4: pd.Timestamp(q['일자']).to_pydatetime(),
                           5: round(q['지급'], 2), 6: round(q['청구'], 2), 8: 비})
            ws.cell(r, 4).number_format = 'yyyy-mm-dd'
        b2 = 새[0] - 1
        새r = 놓기(줄[t2])
        범위고침(새r, s2 + 2 + (새r - t2), t2 - 1 + (새r - t2), a2, b2)
        for c in range(1, 최대열 + 1):          # 합계 식 범위를 새 표에 맞춤
            v = ws.cell(새r, c).value
            if isinstance(v, str) and v.startswith('=SUM('):
                ws.cell(새r, c).value = re.sub(r'([A-Z]+)\d+:([A-Z]+)\d+', rf'\g<1>{a2}:\g<2>{b2}', v)
        r = 제목(f'※ 합계 {선급잔:,.0f} = BS 선급금 ({md(끝날)} 원장) · 지급처별로 먼저 지급한 것부터 청구 · 대체 금액을 지움', 굵게=False)
        # ③ 머리까지 (② 합계 다음 줄 ~ ③ 머리)
        for r in range(t2 + 1, s3 + 2):
            if str(줄[r]['값'].get(2) or '').startswith('※'):
                continue
            놓기(줄[r])
        # ③ 표
        틀3 = 줄[d3[0]]
        a3 = 새[0]
        for d in (행3 or [None]):
            값 = ({c: None for c in (2, 3, 5, 6, 8)} if d is None else
                  {2: d['거래처'], 3: d['계정'], 5: round(d['잔액'], 2), 6: round(d['지급'], 2), 8: d['비고']})
            놓기(틀3, 값)
        b3 = 새[0] - 1
        for r in range(t3, max(줄) + 1):
            새r = 놓기(줄[r])
            if r == t3:
                off = 새r - r
                범위고침(새r, d3[0] + off, d3[1] + off, a3, b3)

        # ② 표 지급일이 숫자(46198)로 보이던 것 → 날짜 모양
        머2 = next((r for r in range(1, ws.max_row + 1) if ws.cell(r, 4).value == '지급일'), None)
        if 머2:
            for r in range(머2 + 1, ws.max_row + 1):
                if str(ws.cell(r, 2).value or '').strip() == '합계':
                    break
                if isinstance(ws.cell(r, 4).value, (int, float)) and 30000 < ws.cell(r, 4).value < 80000:
                    ws.cell(r, 4).number_format = 'yyyy-mm-dd'

    # ── 머리 글 · 기준일 · 메모
    for r in range(1, ws.max_row + 1):
        for c in range(2, 최대열 + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str) or v.startswith('='):
                continue
            v2 = re.sub(r'\d{1,2}/\d{1,2}까지 회수', f'{md(끝날)}까지 회수', v)
            v2 = re.sub(r'\d{1,2}/\d{1,2} 장부잔액', f'{md(끝날)} 장부잔액', v2)
            v2 = re.sub(r'\d{1,2}/\d{1,2} 이후 (회수|지급)', lambda m: f'{md(다음날)} 이후 {m.group(1)}', v2)
            if v2 != v:
                ws.cell(r, c).value = v2
    if ws['B3'].value: ws['B3'] = (f'단위 USD · 장부 = 퀵북 원장 {md(끝날)} 결산 · 결산 후 회수 · 지급 = 원장 {md(원장끝)}까지'
                + (f' + 통장 {md(통장끝)}까지 (2026년 금융거래 파일)' if 통장끝 is not None else ''))
    기준칸 = next(((r, c) for r in range(1, 6) for c in range(2, 최대열 + 1) if ws.cell(r, c).value == '기준일'), None)
    if 기준칸:
        ws.cell(기준칸[0], 기준칸[1] + 1).value = pd.Timestamp(기준일).to_pydatetime()
    메모3 = b3 + 2
    if 메모3 <= ws.max_row and str(ws.cell(메모3, 2).value or '').startswith('※'):
        ws.cell(메모3, 2).value = (f'※ {md(다음날)} 이후 지급 = 원장({md(원장끝)}까지) 차변'
                                   + (f' + 통장({md(통장끝)}까지) 출금' if 통장끝 is not None else '')
                                   + ' · 결산월 말 잔액을 넘는 지급은 넣지 않았습니다.')
    안내.append(f'채권채무 : ① 매출채권 · 선수금 집계와 팩토링({", ".join(팩토링이름) or "없음"}) 세부 원장, ② 선급금 {len(선급행)}줄, ③ 채무 {len(행3)}줄을 새로 썼습니다.'
              + (f' 미지급금 차변잔액(결산조정분개로 자산 대체)은 채무 표에서 뺐습니다 : {" · ".join(건너뜀)}' if 건너뜀 else ''))
    return 안내


커_적요별칭 = [tuple(x) for x in _비밀값('ctk_memo_alias', [])]   # 거래처가 빈 매출 줄을 적요 글자로 찾을 때 (Secrets · 앞에 있는 것부터)


def _커_이름채우기(g, W):
    """거래처가 빈 원장 줄 채우기 : ① 전월 패키지 원장(같은 전표 · 계정 · 금액) ② 같은 전표의 다른 줄
    ③ (매출 · 매출채권 · 선수금 줄만) 적요 글자 → 원장에 있는 거래처 이름."""
    g = g.copy()
    빈 = g['거래처원'].isna()
    빈줄 = int(빈.sum())
    이전 = {}
    for r in range(6, W.max_row + 1):
        b, n = W.cell(r, 2).value, W.cell(r, 5).value
        if not b or b == 'Beginning Balance' or W.cell(r, 3).value is None or not n:
            continue
        이전[(_커_번호(W.cell(r, 6).value), str(b).strip(), round(float(W.cell(r, 9).value or 0), 2),
              round(float(W.cell(r, 10).value or 0), 2))] = n
    키 = [(_커_번호(t), a, round(float(dr), 2), round(float(cr), 2))
          for t, a, dr, cr in zip(g['TID'], g['계정'], g['차'], g['대'])]
    g['거래처원'] = [n if n is not None and not (isinstance(n, float) and pd.isna(n)) else 이전.get(k)
                     for n, k in zip(g['거래처원'], 키)]
    # ② 같은 전표 안에 거래처가 하나뿐이면 그 이름
    하나 = g.dropna(subset=['거래처원']).groupby(g['TID'].map(_커_번호))['거래처원'].agg(lambda s: s.iloc[0] if s.nunique() == 1 else None)
    빈2 = g['거래처원'].isna()
    g.loc[빈2, '거래처원'] = g.loc[빈2, 'TID'].map(_커_번호).map(하나)
    # ③ 적요 글자
    이름들 = sorted({str(n) for n in list(g['거래처원'].dropna()) + list(이전.values())}, key=len)
    def 찾기(글):
        u = str(글).upper()
        for k, 목표 in 커_적요별칭:
            if k.upper() in u:
                후보 = sorted([n for n in 이름들 if 목표.upper() in n.upper()],
                            key=lambda n: (0 if 'CUSTOMER' in n.upper() else 1, len(n)))
                if 후보:
                    return 후보[0]
        return None
    대상 = g['거래처원'].isna() & g['계정'].str.match(r'^(11501|511|21511|11803|536)')
    g.loc[대상, '거래처원'] = g.loc[대상, '적요'].map(찾기)
    g['거래처원'] = g['거래처원'].where(g['거래처원'].notna(), None)
    g['고객'] = g['고객'].where(g['고객'].notna(), g['거래처원'])
    g['공급'] = g['공급'].where(g['공급'].notna(), g['거래처원'])
    남음 = int(g['거래처원'].isna().sum())
    return g, {'빈줄': 빈줄, '채움': 빈줄 - 남음, '남음': 남음}


def 커_패키지만들기(원장바이트, 전월바이트, 결산월=None, 월말환율=None, 평균환율=None, 통장바이트=None):
    """돌려주는 값 : (엑셀 bytes, 안내 dict)"""
    g = 커_원장읽기(원장바이트)
    통장 = 커_통장읽기(통장바이트) if 통장바이트 else None
    연도 = int(g['일자'].dt.year.mode().iat[0])
    if 결산월 is None:
        jm = g[g.get('Transaction type', pd.Series(dtype=str)).astype(str).str.contains('Journal', na=False)]['일자']
        결산월 = int((jm if len(jm) else g['일자']).max().month)
    끝날 = pd.Timestamp(연도, 결산월, calendar.monthrange(연도, 결산월)[1])
    wb = openpyxl.load_workbook(io.BytesIO(전월바이트))
    if '26년 원장_raw' not in wb.sheetnames:
        raise ValueError('전월 연결패키지에 「26년 원장_raw」 시트가 없습니다.')
    매핑 = {}
    if 'BS_IS_매핑' in wb.sheetnames:
        M = wb['BS_IS_매핑']
        for r in range(2, M.max_row + 1):
            a = M.cell(r, 1).value
            if a:
                매핑[str(a).strip().lower()] = (str(M.cell(r, 2).value or '').strip(), str(M.cell(r, 3).value or '').strip())
    안내 = {'연도': 연도, '결산월': 결산월, '원장끝': g['일자'].max(), '통장끝': (통장['일자'].max() if 통장 is not None else None),
            '알림': []}
    빠진계정 = sorted({a for a in g[g['일자'] <= 끝날]['계정'].unique() if a.lower() not in 매핑})
    if 빠진계정 and 'BS_IS_매핑' in wb.sheetnames:
        과목 = {}
        for 시트 in ('BS', 'IS'):
            if 시트 in wb.sheetnames:
                for r in range(8, wb[시트].max_row + 1):
                    v = wb[시트].cell(r, 5).value
                    if v:
                        과목.setdefault(str(v).strip(), 시트)
        M = wb['BS_IS_매핑']
        끝줄 = max(r for r in range(1, M.max_row + 1) if M.cell(r, 1).value)
        넣음 = []
        for a in 빠진계정:
            한 = re.sub(r'\s*\(.*\)\s*$', '', re.sub(r'^\d+\s*', '', a)).strip()
            # 연결패키지 IS 는 외환 · 환산을 「…손익」 하나로 모아 차익 · 차손 줄에 나눕니다
            한 = {'외환차손': '외환차손익', '외환차익': '외환차손익',
                  '외화환산손실': '외화환산손익', '외화환산이익': '외화환산손익'}.get(한, 한)
            if 한 in ('외환차손익', '외화환산손익'):
                과목.setdefault(한, 'IS')
            if 한 in 과목:
                끝줄 += 1
                for c, v in enumerate((a, 과목[한], 한), 1):
                    M.cell(끝줄, c, v)
                    _커_입히기(M.cell(끝줄, c), _커_스타일(M.cell(끝줄 - 1, c)))
                매핑[a.lower()] = (과목[한], 한)
                넣음.append(f'{a} → {과목[한]} {한}')
        if 넣음:
            안내['알림'].append('BS_IS_매핑 에 새 계정을 넣었습니다 (계정 이름과 같은 과목) : ' + ' · '.join(넣음))
        빠진계정 = [a for a in 빠진계정 if a.lower() not in 매핑]
    if 빠진계정:
        안내['알림'].append('BS_IS_매핑 에 없는 계정 : ' + ' · '.join(빠진계정) + ' — 매핑 시트에 넣어야 BS · IS 에 잡힙니다.')

    # 전월 결산월 (BS B2 의 날짜)
    전월 = None
    if 'BS' in wb.sheetnames:
        m = re.search(r'(\d{4})년\s*(\d{1,2})월', str(wb['BS']['B2'].value or ''))
        전월 = int(m.group(2)) if m else None
    안내['전월'] = 전월

    # ⓪ 거래처 이름 채우기 (원장 양식에 따라 분개 줄에 거래처가 비어 있음)
    W = wb['26년 원장_raw']
    g, 안내['이름채움'] = _커_이름채우기(g, W)
    if 안내['이름채움'].get('남음'):
        안내['알림'].append(f'원장에 거래처 이름이 빈 줄 {안내["이름채움"]["빈줄"]:,}건 중 {안내["이름채움"]["채움"]:,}건을 '
                            f'전월 패키지 · 같은 전표 · 적요로 채웠고, {안내["이름채움"]["남음"]:,}건은 비어 있습니다 '
                            '(퀵북에서 Customer · Vendor 열이 있는 양식으로 내려받으면 다 채워집니다).')

    # ① 원장 시트
    새분류, 건수 = _커_원장시트(W, g, 끝날, 매핑)
    W['A3'] = f'{연도}.01.01 ~ {끝날:%Y.%m.%d}'
    안내['원장건수'] = 건수
    안내['새분류'] = 새분류

    # ② BS · IS 제목 · 환율
    말일 = 끝날.day
    if 'BS' in wb.sheetnames:
        B = wb['BS']
        B['B2'] = re.sub(r'\d{4}년\s*\d{1,2}월\s*\d{1,2}일', f'{연도}년 {결산월:02d}월 {말일}일', str(B['B2'].value or '')) or \
            f'제1기 (당) : {연도}년 {결산월:02d}월 {말일}일 현재'
        B['L5'] = f'{연도 % 100}년 {결산월}월말 환율'
        if 월말환율:
            B['M5'] = float(월말환율)
    if 'IS' in wb.sheetnames:
        I = wb['IS']
        I['B2'] = re.sub(r'(\d{4}\.01\.01~)\d{4}\.\d{2}\.\d{2}', rf'\g<1>{연도}.{결산월:02d}.{말일}', str(I['B2'].value or ''))
        I['L5'] = f"'{연도 % 100}년 평균환율(01월~{결산월}월)"
        if 평균환율:
            I['M5'] = float(평균환율)
    if not 월말환율 or not 평균환율:
        안내['알림'].append('환율(BS M5 월말 · IS M5 평균)을 넣지 않아 전월 값 그대로입니다 — KRW 칸 확인 필요.')

    # ③ 「1~8월」 같은 기간 글자 (요약보고 · 월별 실적집계)
    if 전월 and 전월 != 결산월:
        for n in ('요약보고', '월별 실적집계'):
            if n not in wb.sheetnames:
                continue
            for row in wb[n].iter_rows():
                for c in row:
                    v = c.value
                    if isinstance(v, str) and not v.startswith('='):
                        v2 = v.replace(f'1~{전월}월', f'1~{결산월}월').replace(f'{전월}월 누적', f'{결산월}월 누적')
                        if v2 != v:
                            c.value = v2
        안내['알림'].append(f'요약보고 ① · ①-0 · 4(총액) 의 거래액 숫자는 원장만으로 나오지 않아 {전월}월 값 그대로입니다 — '
                            '거래처별 손익 자료로 바꿔 넣어 주세요.')

    # ④ 채권채무
    if '채권채무' in wb.sheetnames:
        try:
            안내['알림'] += _커_채권채무(wb['채권채무'], g, 통장, 끝날, 매핑)
        except Exception as e:
            안내['알림'].append(f'채권채무 시트는 그대로 두었습니다 ({e})')
    if 통장 is None:
        안내['알림'].append('통장 파일이 없어 채권채무의 결산 후 회수 · 지급은 원장 안의 거래만 반영했습니다.')

    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), 안내


커_계좌짝 = [tuple(x) for x in _비밀값('ctk_bank_pairs', [])]     # 원장 현금 계정 이름 ↔ 금융거래 파일 계좌(시트) 이름 (Secrets)


def 커_통장대사(g, 통장):
    """원장의 보통예금 계정 거래와 통장(금융거래 파일 자금data(US)) 거래가 맞는지 봅니다.
    기간 = 두 자료가 모두 있는 날까지. ① 같은 금액 · 3일 안 → 1:1 짝 ② 남은 것은 같은 날 합계로 다시 맞춤.
    돌려주는 값 : [{계좌, 기간끝, 통장{건,입,출}, 원장{건,입,출}, 짝, 통장만[...], 원장만[...]}]"""
    끝 = min(g['일자'].max(), 통장['일자'].max())
    결과 = []
    현금 = g[g['계정'].str.match(r'^1110')]
    for 영, 한 in 커_계좌짝:
        ga = 현금[현금['계정'].str.upper().str.contains(영)]
        ba = 통장[통장['계좌(시트)'].astype(str).str.contains(한) & ~통장['계좌(시트)'].astype(str).str.contains('대출')]
        if ga.empty and ba.empty:
            continue
        ga = ga[ga['일자'] <= 끝]
        ba = ba[ba['일자'] <= 끝]
        G = [{'일자': r['일자'], '금액': round(float(r['차']) - float(r['대']), 2), '적요': r['적요'],
              '거래처': 커_이름정리(r['거래처원']), '짝': False} for _, r in ga.iterrows()]
        B = [{'일자': r['일자'], '금액': round(float(r['입금']) - float(r['출금']), 2), '적요': str(r.get('적요') or ''),
              '짝': False} for _, r in ba.iterrows() if abs(float(r['입금']) - float(r['출금'])) >= 0.005]
        짝 = 0
        for b_ in sorted(B, key=lambda x: x['일자']):          # ① 1:1 (같은 금액 · 3일 안 · 가장 가까운 날)
            후보 = [x for x in G if not x['짝'] and abs(x['금액'] - b_['금액']) < 0.005
                    and abs((x['일자'] - b_['일자']).days) <= 3]
            if 후보:
                x = min(후보, key=lambda x: abs((x['일자'] - b_['일자']).days))
                x['짝'] = b_['짝'] = True
                짝 += 1
        날짜들 = {x['일자'] for x in G + B if not x['짝']}           # ② 같은 날 합계
        for d in 날짜들:
            gs = [x for x in G if not x['짝'] and x['일자'] == d]
            bs = [x for x in B if not x['짝'] and x['일자'] == d]
            if gs and bs and abs(sum(x['금액'] for x in gs) - sum(x['금액'] for x in bs)) < 0.005:
                for x in gs + bs:
                    x['짝'] = True
                짝 += len(bs)
        결과.append({'계좌': ga['계정'].iloc[0] if len(ga) else 한, '기간끝': 끝,
                     '통장': {'건': len(B), '입': sum(x['금액'] for x in B if x['금액'] > 0),
                              '출': -sum(x['금액'] for x in B if x['금액'] < 0)},
                     '원장': {'건': len(G), '입': sum(x['금액'] for x in G if x['금액'] > 0),
                              '출': -sum(x['금액'] for x in G if x['금액'] < 0)},
                     '짝': 짝, '통장만': [x for x in B if not x['짝']], '원장만': [x for x in G if not x['짝']]})
    return 결과


def 실적보고엑셀만들기(원장바이트, 전월바이트):
    """돌려주는 값 : (엑셀 bytes, 안내 정보 dict)"""
    새bio, 전bio = io.BytesIO(원장바이트), io.BytesIO(전월바이트)

    # ── 1. 이번 달 원장 원본
    새 = None
    for s in pd.ExcelFile(새bio).sheet_names:
        for h in (4, 0, 3, 5):
            try:
                d = pd.read_excel(새bio, sheet_name=s, header=h)
            except Exception:
                continue
            if 'Distribution account' in [str(c) for c in d.columns]:
                새 = d
                break
        if 새 is not None:
            break
    if 새 is None:
        raise ValueError('원장 파일에서 「Distribution account」 열을 찾지 못했습니다. '
                         'QuickBooks 에서 내려받은 General Ledger 원본이 맞는지 확인해 주세요.')

    # ── 2. 전월 실적보고자료
    z = zipfile.ZipFile(전bio)
    (올해시트, 올해파일), 시트목록 = _전월원장시트(z)
    전 = pd.read_excel(전bio, sheet_name=올해시트, header=4)

    # ── 3. 손으로 나눈 활동분류·검토메모 이어받기
    전키, 새키 = _이어받기키(전), _이어받기키(새)
    이어받음, 이어받은수 = {}, 0
    for col, 이름 in 이어받기열.items():
        후보 = _열찾기(전, 이름)
        if 후보 is None:
            이어받음[col] = [None] * len(새)
            continue
        사전 = pd.Series(전[후보].values, index=전키)
        사전 = 사전[~사전.index.duplicated()]
        옮김 = 새키.map(사전)
        이어받음[col] = [None if (v is None or (isinstance(v, float) and pd.isna(v))
                                 or str(v).strip() == '') else str(v) for v in 옮김]
        if col == 'Q':
            이어받은수 = sum(1 for v in 이어받음[col] if v)

    # ── 3-2. 전월에 없던 새 거래는 지난달까지의 분류를 배워서 스스로 채웁니다
    시트칸 = dict(_시트지도(z))
    공유 = _공유글자(z)
    집계시트 = next((n for n in 시트칸 if '월별' in n and '실적집계' in n), None)
    유효세부 = set(_시트글자(z, 시트칸[집계시트], 'E', 공유)) if 집계시트 else set()
    매핑시트 = next((n for n in 시트칸 if 'BS_IS_매핑' in n), None)
    매핑표 = _매핑표읽기(z, 시트칸[매핑시트], 공유) if 매핑시트 else {}
    매핑분류 = {k: v[0] for k, v in 매핑표.items()}
    분류센것 = _자동분류(새, 이어받음, _분류학습(전), 유효세부, 매핑분류)

    날짜들 = [d for d in (_날짜일련(v) for v in 새['Transaction date']) if d is not None]
    처음 = 엑셀기준일 + datetime.timedelta(days=min(날짜들))
    마지막 = 엑셀기준일 + datetime.timedelta(days=max(날짜들))
    기간글 = f'{처음:%Y.%m.%d} ~ {마지막:%Y.%m.%d}'
    보고달 = 마지막.month

    # ── 4. 원장 시트를 갈아끼웁니다
    새xml = _원장시트XML(z.read(올해파일).decode('utf-8'), 새, 이어받음, 기간글, 매핑표)
    끝행 = 5 + len(새)
    # 「당월 실적집계」는 이제 쓰지 않는 시트라 결과 파일에서 통째로 뺍니다.
    #   (담당자가 원본에서 지우셨습니다. 옛 파일을 올리셔도 결과에는 들어가지 않습니다.)
    당월시트 = next((n for n in 시트칸 if '당월' in n and '실적집계' in n), None)
    당월xml, 당월파일, 달바뀜, 수식고침 = None, None, False, []
    뺀파트, 뺀패치 = (_시트빼기(z, 당월시트) if 당월시트 else (set(), {}))
    뺀시트 = 당월시트 if 뺀파트 else None

    # ── 4-1. 전월 자료가 「연결패키지」(BS·IS 시트가 있는 파일)이면
    #        BS·IS 의 당기 제목·환율 라벨만 이번 달로 바꿉니다. (월별 실적집계는 원장_raw 수식이라 그대로)
    #        (숫자는 「원장_raw」를 보는 SUMIFS 수식이라 엑셀이 열릴 때 스스로 다시 계산합니다)
    패키지형 = bool(시트칸.get('BS') and 시트칸.get('IS'))   # 월별 실적집계가 함께 있어도 패키지
    패키지글 = {}
    if 패키지형:
        말일 = calendar.monthrange(마지막.year, 보고달)[1]
        패키지글 = {
            시트칸['BS']: {'B2': f'제? (당) 기 : {마지막.year}년 {보고달:02d}월 {말일}일',
                          'L5': f'{마지막.year % 100}년 {보고달}월말 환율'},
            시트칸['IS']: {'B2': f'제?(당)기 : {마지막.year}.01.01~{마지막.year}.{보고달:02d}.{말일}',
                          'L5': f"'{마지막.year % 100}년 평균환율(01월~{보고달}월)"},
        }

    # ── 4-2. 「26년 자금」 시트도 이번 달까지로 다시 씁니다 (연결패키지에는 이 시트가 없어 건너뜀)
    if not any('자금' in str(n) for n in 시트칸):
        자금xml, 새스타일, 자금파일, 자금안내 = None, None, None, {'없음': True}
    else:
        try:
            자금xml, 새스타일, 자금파일, 자금안내 = _자금시트만들기(
                z, 시트칸, 공유, 새, 이어받음, 매핑표, 마지막.year, 보고달)
        except Exception as e:                  # 자금 시트 하나 때문에 전체가 막히면 안 됩니다
            자금xml, 새스타일, 자금파일, 자금안내 = None, None, None, {'오류': str(e)}
    buf = io.BytesIO()
    out = zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=6)
    for it in z.infolist():
        이름 = it.filename
        if 이름 == 'xl/calcChain.xml':      # 낡은 계산순서표는 빼야 「복구됨」이 안 뜹니다
            continue
        if 이름 in 뺀파트:                  # 「당월 실적집계」와 거기에만 딸린 파트들
            continue
        데이터 = 뺀패치.get(이름) or z.read(이름)
        if 이름 == 올해파일:
            데이터 = 새xml.encode('utf-8')
        elif 당월xml is not None and 이름 == 당월파일:
            데이터 = 당월xml.encode('utf-8')
        elif 자금xml is not None and 이름 == 자금파일:
            데이터 = 자금xml.encode('utf-8')
        elif 새스타일 is not None and 이름 == 'xl/styles.xml':
            데이터 = 새스타일.encode('utf-8')
        elif 이름 in 패키지글:                 # 연결패키지 BS·IS 제목 · 환율 라벨
            데이터 = 패_글자칸(데이터.decode('utf-8'), 패키지글[이름]).encode('utf-8')
        elif 이름 == 'xl/workbook.xml':
            데이터 = re.sub(r'<calcPr[^>]*/>',
                            '<calcPr calcId="191029" fullCalcOnLoad="1"/>',
                            데이터.decode('utf-8'), count=1).encode('utf-8')
        elif 이름 == '[Content_Types].xml':
            데이터 = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '',
                            데이터.decode('utf-8')).encode('utf-8')
        elif 이름 == 'xl/_rels/workbook.xml.rels':
            데이터 = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '',
                            데이터.decode('utf-8')).encode('utf-8')
        elif 'pivotCacheDefinition' in 이름 and 이름.endswith('.xml'):
            t = 데이터.decode('utf-8')
            t = re.sub(r'(<pivotCacheDefinition\b[^>]*?)(\s*>)',
                       lambda m: re.sub(r'\s+refreshOnLoad="[^"]*"', '', m.group(1))
                       + ' refreshOnLoad="1"' + m.group(2), t, count=1)
            t = re.sub(r'(<worksheetSource ref="[A-Z]+5:[A-Z]+)\d+(")',
                       rf'\g<1>{max(끝행 + 3000, 30000)}\g<2>', t)
            데이터 = t.encode('utf-8')
        out.writestr(it, 데이터)
    out.close()

    연도, 월끝 = 마지막.year, 마지막.month
    if 뺀시트:                       # 뺀 시트는 결과 파일의 시트 목록에서도 지웁니다
        시트목록 = [n for n in 시트목록 if n != 뺀시트]
    거래 = int(새['Distribution account'].notna().sum())
    자동합 = sum(분류센것.get(k, 0) for k in
                 ('적요일치', '거래처일치', '계정과목대표값', '계정과목최빈값'))
    안내 = {'연도': 연도, '월끝': 월끝, '기간': 기간글, '올해시트': 올해시트,
            '행수': len(새), '거래건수': 거래, '이어받음': 이어받은수,
            '새거래': int(새키.isin(set(전키)).eq(False).sum()),
            '시트수': len(시트목록), '시트목록': 시트목록,
            '당월시트': 당월시트, '기준월맞춤': 달바뀜, '보고달': 보고달,
            '뺀시트': 뺀시트, '패키지형': 패키지형,
            '수식고침': 수식고침, '자금': 자금안내,
            '자동분류': 자동합, '분류내역': 분류센것,
            '확인필요': 분류센것.get('확인필요', 0) + 분류센것.get('신규계정', 0),
            '집계밖': 분류센것.get('집계밖', 0), '매핑밖': 분류센것.get('매핑밖', 0),
            '되살린열': list(_되살린열)}
    결과 = buf.getvalue()
    # ── 5. 연결패키지면 ① 원장 파일의 퀵북 BS·IS 를 「퀵북 BS」「퀵북 IS」 에 넣고 「검증」 시트로 비교합니다
    if 패키지형:
        try:
            결과 = 과제잠그기(결과, (마지막.year, 보고달))      # 개선과제 · 판관비절감 : 지난 달 칸 잠금
        except Exception:
            pass
        try:
            결과, 안내['퀵북'] = 검_패키지에넣기(결과, 검_퀵북표읽기(원장바이트))
        except Exception as e:             # 검증 시트 하나 때문에 전체가 막히면 안 됩니다
            안내['퀵북'] = '넣지 못했습니다 — ' + str(e)[:120]
    return 결과, 안내


@st.cache_data(max_entries=1, show_spinner='실적보고 엑셀을 만드는 중입니다...')
def 실적보고엑셀캐시(원장바이트, 전월바이트):
    """같은 파일이면 한 번만 만듭니다 (내려받기를 눌러도 다시 안 만듦)."""
    return 실적보고엑셀만들기(원장바이트, 전월바이트)


자료실 = '자료실'


def 오늘날짜():
    """보고 계신 분의 시계로 「오늘」 — 웹서버는 영국 시간이라 하루가 어긋납니다.

    ※ 한국에서 8월 29일 아침이어도 서버는 아직 8월 28일이라, 파일 이름이
      하루 뒤처져 붙던 것을 막습니다. 브라우저가 알려 주는 시간대를 씁니다.
    """
    try:
        이름 = getattr(st.context, 'timezone', None)
        if 이름:
            import zoneinfo
            return datetime.datetime.now(zoneinfo.ZoneInfo(str(이름))).date()
    except Exception:
        pass
    try:
        분 = getattr(st.context, 'timezone_offset', None)   # 한국이면 −540
        if 분 is not None:
            return (datetime.datetime.now(datetime.timezone.utc)
                    - datetime.timedelta(minutes=int(분))).date()
    except Exception:
        pass
    return datetime.date.today()
# 파일 이름 끝의 「2026.08.27_v1」 을 「만든 날짜 + 그날의 몇 번째」로 봅니다
날짜버전패턴 = re.compile(r'(\d{4})\.(\d{1,2})\.(\d{1,2})[_\s]*v\s*(\d+)\s*\.xlsx$', re.I)
# 예전 이름(…클로드작성_v24.xlsx)도 목록에는 그대로 보이게 합니다
버전패턴 = re.compile(r'v\s*(\d+)\s*\.xlsx$', re.I)
# 파일 이름 앞쪽의 「26년 6월」은 그 파일이 다루는 실적 월입니다
연월패턴 = re.compile(r'(\d{2})\s*년\s*(\d{1,2})\s*월')


def _이름속연월(이름):
    m = 연월패턴.search(str(이름) or '')
    return (int(m.group(1)), int(m.group(2))) if m else None


def _이름속날짜(이름):
    m = 날짜버전패턴.search(str(이름) or '')
    if not m:
        return None
    try:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def 자료실목록():
    """자료실 폴더에 쌓인 실적보고 엑셀을 새 것부터 돌려줍니다."""
    try:
        이름들 = [n for n in os.listdir(자료실) if n.lower().endswith('.xlsx')
                  and not n.startswith('~$') and 버전패턴.search(n)
                  and '연결패키지' not in n]
    except OSError:
        return []
    목록 = []
    for n in 이름들:
        경로 = os.path.join(자료실, n)
        try:
            st_ = os.stat(경로)
        except OSError:
            continue
        목록.append({'이름': n, '경로': 경로, '버전': int(버전패턴.search(n).group(1)),
                     '연월': _이름속연월(n), '만든날': _이름속날짜(n),
                     '시각': st_.st_mtime, '크기': st_.st_size})
    목록.sort(key=lambda d: (d['시각'], d['버전']), reverse=True)
    return 목록


def 다음버전(오늘=None, 처음부터=False):
    """번호는 「만든 날」로 셉니다.

    ※ 같은 날 또 만들면 v2 · v3 … 이고, 날이 바뀌면 다시 v1 부터입니다.
      (8월 27일에 두 번 만들면 2026.08.27_v1 · 2026.08.27_v2,
       28일에 만들면 2026.08.28_v1)
    """
    if 처음부터:
        return 1
    오늘 = 오늘 or 오늘날짜()
    번호 = [d['버전'] for d in 자료실목록() if d['만든날'] == 오늘]
    return (max(번호) + 1) if 번호 else 1


def 보고서파일이름(연도, 월끝, 버전, 오늘=None):
    오늘 = 오늘 or 오늘날짜()
    return (f'CTKonnect실적분석_{연도 % 100}년 {월끝}월 누적_클로드작성_'
            f'{오늘:%Y.%m.%d}_v{버전}.xlsx')


def _달표시(d):
    """자료실 목록에 「6월 / 7월」처럼 그 파일이 다루는 실적 월을 보여 줍니다."""
    return f'{d["연월"][1]}월' if d.get('연월') else '—'


def _만든날표시(d):
    return f'{d["만든날"]:%Y.%m.%d}' if d.get('만든날') else '—'

# ══════════════════════════════════════════════════════════════
# 10. 연결패키지 만들기 — 원장(퀵북 BS·IS) + 지난번 패키지 양식 → 일곱 장짜리 엑셀
#     BS · IS · CE · MC · CF · CF정산서 · 결산조정분개 (수식 없이 값만)
# ══════════════════════════════════════════════════════════════
def 패_풀기(s):
    return _html.unescape(s)


def 패_엑스(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))


def 패_공유(z):
    try:
        sx = z.read('xl/sharedStrings.xml').decode('utf-8')
    except KeyError:
        return []
    return [패_풀기(re.sub(r'<[^>]+>', '', m.group(1)))
            for m in re.finditer(r'<si>(.*?)</si>', sx, re.S)]


def 패_시트지도(z):
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rel = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid = _관계사전(rel)
    지도 = {}
    for m in re.finditer(r'<sheet\b([^>]*?)/>', wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        지도[패_풀기(str(a.get('name')))] = (rid.get(a.get('r:id'), 'xl/'), a.get('r:id'))
    return 지도


def 패_칸읽기(z, 파일, 공유):
    """{행: {열: ('값'|'수식', 내용)}}

    ※ 엑셀은 같은 모양의 수식을 「공유 수식」으로 한 번만 적어 두고 나머지 칸은
      번호(si)만 남깁니다. 그대로 읽으면 수식이 빈 글자로 보여서, 대표 수식을 찾아
      줄 번호만 바꿔 채워 넣습니다.
    """
    x = z.read(파일).decode('utf-8')
    if '<sheetData>' not in x:
        return {}
    본문 = x.split('<sheetData>', 1)[1].rsplit('</sheetData>', 1)[0]
    대표 = {}
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"[^>]*?[^/]>\s*<f([^>]*?[^/])>(.*?)</f>',
                         본문, re.S):
        si = re.search(r'si="(\d+)"', m.group(3))
        if si and 'shared' in m.group(3) and m.group(4).strip():
            대표.setdefault(si.group(1), (int(m.group(2)), 패_풀기(m.group(4))))

    def _옮기기(글, 원행, 새행):
        if 원행 == 새행:
            return 글
        return re.sub(r'(\$?[A-Z]{1,3}\$?)' + str(원행) + r'(?![0-9])',
                      lambda mm: mm.group(1) + str(새행)
                      if '$' not in mm.group(1)[-1:] else mm.group(0), 글)

    표 = {}
    for rm in re.finditer(r'<row[^>]*?\br="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', 본문, re.S):
        r = int(rm.group(1))
        칸 = {}
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)',
                              rm.group(2) or '', re.S):
            c, 속, 몸 = cm.group(1), cm.group(2), cm.group(3) or ''
            fm = re.search(r'<f([^>]*?)(?:/>|>(.*?)</f>)', 몸, re.S)
            if fm:
                글 = 패_풀기(fm.group(2)) if fm.group(2) else ''
                if not 글.strip():
                    si = re.search(r'si="(\d+)"', fm.group(1))
                    if si and si.group(1) in 대표:
                        원행, 본 = 대표[si.group(1)]
                        글 = _옮기기(본, 원행, r)
                캐시 = re.search(r'<v>(.*?)</v>', 몸, re.S)
                칸[c] = ('수식', 글, 패_풀기(캐시.group(1)) if 캐시 else None)
                continue
            vm = re.search(r'<v>(.*?)</v>', 몸, re.S)
            if vm:
                v = vm.group(1)
                칸[c] = ('값', 공유[int(v)] if ' t="s"' in 속 and v.isdigit()
                         and int(v) < len(공유) else 패_풀기(v), None)
            elif '<is>' in 몸:
                칸[c] = ('값', 패_풀기(re.sub(r'<[^>]+>', '', 몸)), None)
        if 칸:
            표[r] = 칸
    return 표


def 패_숫자(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def 패_참조풀기(수식, 열, 내행, 있는행):
    """소계 수식을 [(행, +1/−1)] 로 풉니다. SUM(a:b) · SUM(a,b) · +A−B 를 다룹니다."""
    s = 수식.lstrip('=').replace(' ', '')
    쓸 = []

    def _덩어리(m):
        부호 = -1 if m.group(1) == '-' else 1
        속 = m.group(2)
        for a, b in re.findall(rf'{열}(\d+):{열}(\d+)', 속):
            쓸.extend((r, 부호) for r in range(int(a), int(b) + 1))
        for r in re.findall(rf'(?<![A-Z]){열}(\d+)',
                            re.sub(rf'{열}\d+:{열}\d+', '', 속)):
            쓸.append((int(r), 부호))
        return '#'

    남 = re.sub(rf'([+-]?)SUM\(([^()]*)\)', _덩어리, s)
    for m in re.finditer(rf'([+-]?)(?<![A-Z]){열}(\d+)', 남):
        쓸.append((int(m.group(2)), -1 if m.group(1) == '-' else 1))
    본, 본행 = [], set()
    for r, g in 쓸:
        if r != 내행 and r in 있는행 and r not in 본행:
            본.append((r, g))
            본행.add(r)
    return 본 or None



# ── 작은 수식 계산기 ────────────────────────────────────────────
#    이 파일의 BS·IS 시트는 SUM · ROUND · IFERROR · 더하기빼기만 씁니다.
#    (SUMIF 는 「계정과목으로 금액을 찾아오는 줄」이라 따로 다룹니다)
패_참조 = re.compile(r'\$?([A-Z]{1,3})\$?(\d+)')


def 패_펼치기(속, 잡기):
    쓸 = []

    def _열번호(c):
        n = 0
        for ch in c:
            n = n * 26 + (ord(ch) - 64)
        return n

    def _열이름(n):
        이름 = ''
        while n:
            n, 나 = divmod(n - 1, 26)
            이름 = chr(65 + 나) + 이름
        return 이름

    for a1, r1, a2, r2 in re.findall(r'\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)', 속):
        if a1 == a2:                               # 세로 범위 (C8:C20)
            쓸 += [잡기(f'{a1}{r}') for r in range(int(r1), int(r2) + 1)]
        elif r1 == r2:                             # 가로 범위 (C8:N8)
            쓸 += [잡기(f'{_열이름(n)}{r1}')
                   for n in range(_열번호(a1), _열번호(a2) + 1)]
        else:                                      # 네모 범위
            쓸 += [잡기(f'{_열이름(n)}{r}')
                   for n in range(_열번호(a1), _열번호(a2) + 1)
                   for r in range(int(r1), int(r2) + 1)]
    남 = re.sub(r'\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+', '', 속)
    for c, r in 패_참조.findall(남):
        쓸.append(잡기(f'{c}{r}'))
    return '(' + '+'.join(f'({v!r})' for v in 쓸) + ')' if 쓸 else '0'


def 패_계산(수식, 잡기):
    """수식 글자를 숫자로 바꿉니다. 못 풀면 None."""
    F = 수식.lstrip('=').replace(' ', '')
    if 'SUMIF' in F or 'OFFSET' in F or 'INDEX' in F:
        return None
    for _ in range(6):
        새 = re.sub(r'IFERROR\(([^()]*),[^()]*\)', r'(\1)', F)
        새 = re.sub(r'ROUND\(([^()]*),(-?\d+)\)', r'_r((\1),\2)', 새)
        새 = re.sub(r'SUM\(([^()]*)\)', lambda m: 패_펼치기(m.group(1), 잡기), 새)
        새 = re.sub(r'MAX\(([^()]*)\)', r'_mx(\1)', 새)
        새 = re.sub(r'MIN\(([^()]*)\)', r'_mn(\1)', 새)
        새 = re.sub(r'ABS\(([^()]*)\)', r'_ab(\1)', 새)
        if 새 == F:
            break
        F = 새
    if re.search(r'[A-Z]{2,}\(', F):          # 아직 모르는 함수가 남아 있으면 포기
        return None
    F = 패_참조.sub(lambda m: f'({잡기(m.group(1) + m.group(2))!r})', F)
    F = F.replace('=', '==').replace('<==', '<=').replace('>==', '>=').replace('!==', '!=')
    try:
        v = eval(F, {'__builtins__': {}},
                 {'_r': lambda x, n: round(x, n), '_mx': max, '_mn': min, '_ab': abs})
        return float(v)
    except Exception:
        return None


def 패_시트계산(칸들, 덮개=None):
    """시트 한 장의 모든 칸 값을 구합니다. 덮개 = {칸: 값} 은 그 값을 먼저 씁니다."""
    덮개 = 덮개 or {}
    메모, 진행 = {}, set()

    def 잡기(칸):
        if 칸 in 덮개:
            return float(덮개[칸])
        if 칸 in 메모:
            return 메모[칸]
        if 칸 in 진행:                      # 서로 물고 도는 수식은 0 으로 끊습니다
            return 0.0
        m = re.fullmatch(r'([A-Z]{1,3})(\d+)', 칸)
        if not m:
            return 0.0
        셀 = 칸들.get(int(m.group(2)), {}).get(m.group(1))
        if 셀 is None:
            메모[칸] = 0.0
            return 0.0
        if 셀[0] == '값':
            메모[칸] = 패_숫자(셀[1]) or 0.0
            return 메모[칸]
        진행.add(칸)
        v = 패_계산(셀[1], 잡기)
        진행.discard(칸)
        if v is None:                     # 못 푸는 수식은 엑셀이 계산해 둔 값을 씁니다
            v = 패_숫자(셀[2] if len(셀) > 2 else None)
        메모[칸] = 0.0 if v is None else v
        return 메모[칸]

    return 잡기


def 패_퀵북잔액(xls, 시트):
    """퀵북 BS/IS 시트 (A열 계정 · B열 금액) → {계정: 금액}. Total 줄은 뺍니다."""
    d = pd.read_excel(xls, sheet_name=시트, header=None)
    out = {}
    for _i, 줄 in d.iterrows():
        n = 줄.iloc[0]
        if n is None or (isinstance(n, float) and pd.isna(n)):
            continue
        n = str(n).strip()
        if not n or n.lower().startswith('total'):
            continue
        v = pd.to_numeric(pd.Series([줄.iloc[1] if len(줄) > 1 else None]),
                          errors='coerce').fillna(0).iat[0]
        out[n] = out.get(n, 0.0) + float(v)
    return out


# 퀵북 계정 이름이 패키지와 다른 것 (패키지 이름 → 퀵북 이름)
패_계정바꿔치기 = {'Prepayment': 'General Deposit',
                'LOC - LT': 'Loan payable (Long-term)',
                'CITI Bank LOC': 'Citi Bank LOC'}
# 자산 자리에 있지만 실제로는 부채라 퀵북 부호를 뒤집어야 표가 맞는 계정
패_부호뒤집기 = {'GST/HST Payable'}
# 퀵북 한 계정이 패키지에서 이익 줄·손실 줄로 갈라지는 것
패_FX갈라짐 = {'Exchange Gain or Loss':
            ('Gains Foreign Currency Transactions',
             'Losses Foreign Currency Transactions'),
            'Exchange Translation Gain or Loss':
            ('Gain on foreign exchagne translations',
             'Losses on foreign exchagne translations')}
패_남길시트 = ['BS', 'IS', 'CE', 'MC', 'CF', 'CF정산서', '결산조정분개']
패_IS월열 = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
# 맨 앞에 새로 만들어 붙이는 영문 시트
패_영문시트 = ['BS(영문)', 'IS(영문)']
패_꼬리줄 = re.compile(r'(accrual|cash)\s+basis|^\s*\w+day,\s', re.I)


def 패_이름맞추기(n):
    """「Total for Income」 과 「Total Income」 을 같은 줄로 봅니다."""
    n = re.sub(r'\s+', ' ', str(n or '')).strip().lower()
    return re.sub(r'^total for\s+', 'total ', n)


def 패_전기표(z, 지도, 공유, 시트, 열):
    """양식의 전기(작년) 칸을 계정 이름으로 찾아 쓸 수 있게 모읍니다."""
    if 시트 not in 지도:
        return {}
    칸들 = 패_칸읽기(z, 지도[시트][0], 공유)
    잡기 = 패_시트계산(칸들)
    표 = {}
    for r in sorted(칸들):
        a = str(칸들[r].get('A', ('값', ''))[1]).strip()
        if not a or 열 not in 칸들[r]:
            continue
        키 = 패_이름맞추기(a)
        if 키 not in 표:
            표[키] = 잡기(f'{열}{r}')
    return 표


def 패_영문줄(xls, 시트, 전기표):
    """퀵북 시트를 차례 그대로 읽어 (제목들, [(계정, 전기, 당기, 굵게)]) 로."""
    d = pd.read_excel(xls, sheet_name=시트, header=None)
    제목, 줄들 = [], []
    for i, 행 in d.iterrows():
        n = 행.iloc[0]
        n = '' if n is None or (isinstance(n, float) and pd.isna(n)) else str(n).strip()
        v = pd.to_numeric(pd.Series([행.iloc[1] if len(행) > 1 else None]),
                          errors='coerce').iat[0]
        if i < 5:
            if n and n.lower() != 'total':
                제목.append(n)
            continue
        if not n or 패_꼬리줄.search(n):
            continue
        키 = 패_이름맞추기(n)
        당 = None if pd.isna(v) else float(v)
        # 금액이 없는 줄은 구역 머리글이라 전기도 비워 둡니다
        줄들.append((n, None if 당 is None else 전기표.get(키), 당,
                     bool(re.match(r'^total\b', 키))))
    while len(제목) < 3:
        제목.append('')
    return 제목, 줄들


def 패_스타일더하기(sx):
    """서식은 반드시 「맨 뒤」에 덧붙입니다 (앞에 끼우면 번호가 밀려 파일이 깨집니다)."""
    번호 = 900
    while f'numFmtId="{번호}"' in sx:
        번호 += 1
    금액꼴 = f'<numFmt numFmtId="{번호}" formatCode="#,##0.00;\\(#,##0.00\\)"/>'
    if '<numFmts' in sx:
        sx = re.sub(r'<numFmts count="(\d+)">',
                    lambda m: f'<numFmts count="{int(m.group(1)) + 1}">', sx, count=1)
        sx = sx.replace('</numFmts>', 금액꼴 + '</numFmts>', 1)
    else:
        sx = sx.replace('<fonts', f'<numFmts count="1">{금액꼴}</numFmts><fonts', 1)

    m = re.search(r'<fonts count="(\d+)"[^>]*>(<font>.*?</font>)', sx, re.S)
    글꼴수 = int(m.group(1))
    굵은글꼴 = 글꼴수
    첫글꼴 = m.group(2)
    sx = re.sub(r'<fonts count="(\d+)"',
                lambda mm: f'<fonts count="{int(mm.group(1)) + 1}"', sx, count=1)
    sx = sx.replace('</fonts>', 첫글꼴.replace('<font>', '<font><b/>', 1) + '</fonts>', 1)

    m = re.search(r'<borders count="(\d+)"', sx)
    밑줄 = int(m.group(1))
    sx = re.sub(r'<borders count="(\d+)"',
                lambda mm: f'<borders count="{int(mm.group(1)) + 1}"', sx, count=1)
    sx = sx.replace('</borders>',
                    '<border><left/><right/><top/>'
                    '<bottom style="thin"><color indexed="64"/></bottom>'
                    '<diagonal/></border></borders>', 1)

    m = re.search(r'<cellXfs count="(\d+)"', sx)
    첫칸 = int(m.group(1))
    새것 = [
        f'<xf numFmtId="0" fontId="{굵은글꼴}" fillId="0" borderId="0" xfId="0" '
        f'applyFont="1"/>',                                                   # 제목
        f'<xf numFmtId="0" fontId="{굵은글꼴}" fillId="0" borderId="{밑줄}" xfId="0" '
        f'applyFont="1" applyBorder="1" applyAlignment="1">'
        f'<alignment horizontal="center"/></xf>',                             # 머리글
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',     # 글자
        f'<xf numFmtId="{번호}" fontId="0" fillId="0" borderId="0" xfId="0" '
        f'applyNumberFormat="1"/>',                                           # 숫자
        f'<xf numFmtId="0" fontId="{굵은글꼴}" fillId="0" borderId="0" xfId="0" '
        f'applyFont="1"/>',                                                   # 굵은 글자
        f'<xf numFmtId="{번호}" fontId="{굵은글꼴}" fillId="0" borderId="0" xfId="0" '
        f'applyNumberFormat="1" applyFont="1"/>',                             # 굵은 숫자
    ]
    sx = re.sub(r'<cellXfs count="(\d+)"',
                lambda mm: f'<cellXfs count="{int(mm.group(1)) + len(새것)}"', sx, count=1)
    sx = sx.replace('</cellXfs>', ''.join(새것) + '</cellXfs>', 1)
    칸번호 = dict(zip(('제목', '머리', '글자', '숫자', '굵은글자', '굵은숫자'),
                      range(첫칸, 첫칸 + len(새것))))
    return sx, 칸번호


def 패_영문시트XML(제목, 줄들, 전기이름, 서식):
    """새 워크시트 한 장을 통째로 만들어 냅니다."""
    def 글(칸, s, sty):
        return (f'<c r="{칸}" s="{sty}" t="inlineStr"><is>'
                f'<t xml:space="preserve">{패_엑스(s)}</t></is></c>')

    def 수(칸, v, sty):
        return f'<c r="{칸}" s="{sty}"><v>{round(float(v), 2)!r}</v></c>'

    줄XML = []
    for i, s in enumerate(제목[:3], 1):
        if s:
            줄XML.append(f'<row r="{i}">{글(f"A{i}", s, 서식["제목"])}</row>')
    머리 = [('A', 'Account'), ('B', 전기이름), ('C', 제목[2] or 'Current'),
            ('D', 'Change')]
    줄XML.append('<row r="5">' + ''.join(글(f'{c}5', s, 서식['머리'])
                                        for c, s in 머리) + '</row>')
    r = 6
    for 이름, 전, 당, 굵게 in 줄들:
        칸 = [글(f'A{r}', 이름, 서식['굵은글자' if 굵게 else '글자'])]
        숫자꼴 = 서식['굵은숫자' if 굵게 else '숫자']
        if 전 is not None:
            칸.append(수(f'B{r}', 전, 숫자꼴))
        if 당 is not None:
            칸.append(수(f'C{r}', 당, 숫자꼴))
        if 전 is not None and 당 is not None:
            칸.append(수(f'D{r}', 당 - 전, 숫자꼴))
        줄XML.append(f'<row r="{r}">' + ''.join(칸) + '</row>')
        r += 1
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/'
            'spreadsheetml/2006/main">'
            f'<dimension ref="A1:D{max(r - 1, 6)}"/>'
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="16.5"/>'
            '<cols><col min="1" max="1" width="46" customWidth="1"/>'
            '<col min="2" max="4" width="17" customWidth="1"/></cols>'
            '<sheetData>' + ''.join(줄XML) + '</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
            'header="0.3" footer="0.3"/></worksheet>')


def 패_원장월(xls):
    """원장 파일에서 「○○년 N월(BS)」 같은 시트를 찾아 (월, BS시트, IS시트)."""
    bs = is_ = None
    월 = 0
    for s in xls.sheet_names:
        m = re.search(r'(\d+)\s*월\s*\(\s*(BS|IS)\s*\)', str(s), re.I)
        if not m:
            continue
        월 = max(월, int(m.group(1)))
        if m.group(2).upper() == 'BS':
            bs = s
        else:
            is_ = s
    # 시트 이름이 그냥 「BS」「IS」 인 파일 (2026-09 부터 담당자가 이렇게 보내 줌)
    #   → 달은 시트 머리의 "January-September, 2026" 같은 글에서 읽습니다.
    영문달 = {m: i + 1 for i, m in enumerate(
        ['january', 'february', 'march', 'april', 'may', 'june', 'july',
         'august', 'september', 'october', 'november', 'december'])}
    for s in xls.sheet_names:
        이름 = str(s).strip().upper()
        if 이름 not in ('BS', 'IS') or (bs if 이름 == 'BS' else is_):
            continue
        if 이름 == 'BS':
            bs = s
        else:
            is_ = s
        try:
            머리 = pd.read_excel(xls, sheet_name=s, header=None, nrows=4)
            글 = ' '.join(str(v) for v in 머리.iloc[:, 0].tolist() if pd.notna(v)).lower()
            달들 = [영문달[k] for k in re.findall(r'[a-z]+', 글) if k in 영문달]
            if 달들:
                월 = max(월, max(달들))
        except Exception:
            pass
    return 월, bs, is_


def 패_연결패키지만들기(원장바이트, 양식바이트):
    """돌려주는 값 : (엑셀 bytes, 안내 dict)"""
    xls = pd.ExcelFile(io.BytesIO(원장바이트))
    보고월, BS시트, IS시트 = 패_원장월(xls)
    if not BS시트 or not IS시트:
        raise ValueError('원장 파일에서 「○월(BS)」·「○월(IS)」 시트를 찾지 못했습니다. '
                         '퀵북에서 그 달 재무상태표·손익계산서를 같이 담아 주세요.')
    QB_BS, QB_IS = 패_퀵북잔액(xls, BS시트), 패_퀵북잔액(xls, IS시트)
    소BS = {k.lower(): v for k, v in QB_BS.items()}
    소IS = {k.lower(): v for k, v in QB_IS.items()}

    z = zipfile.ZipFile(io.BytesIO(양식바이트))
    공유 = 패_공유(z)
    지도 = 패_시트지도(z)
    for 이름 in ('BS-당기', 'IS-당기', 'BS', 'IS'):
        if 이름 not in 지도:
            raise ValueError(
                f'패키지 양식에서 「{이름}」 시트를 찾지 못했습니다. '
                '②번에는 클로드가 만들어 준 파일이 아니라, 본사에서 받은 '
                '「연결 패키지」 원본(BS-당기·IS-당기 시트가 든 파일)을 올려 주세요.')
    안내 = {'보고월': 보고월, '확인': [], '없는계정': [], '새계정': []}

    # ── 1. BS-당기 : 퀵북 잔액을 넣고, 계정과목(국문)별로 모읍니다 ──
    bs당 = 패_칸읽기(z, 지도['BS-당기'][0], 공유)
    덮개, 쓴것 = {}, set()
    for r in sorted(bs당):
        칸 = bs당[r]
        a = str(칸.get('A', ('값', ''))[1]).strip()
        D = 칸.get('D')
        if not a or D is None or D[0] == '수식':
            continue
        원 = 패_계정바꿔치기.get(a, a)
        if 원 in QB_BS:
            v = QB_BS[원]
        elif 원.lower() in 소BS:
            v = 소BS[원.lower()]
        else:
            v = 0.0
            if abs(패_숫자(D[1]) or 0) > 0.005:
                안내['없는계정'].append(('BS', a, 패_숫자(D[1])))
        if 원 in 패_부호뒤집기:
            v = -v
            안내['확인'].append(f'BS 「{a}」 는 퀵북이 {-v:,.2f} 인데 자산 자리에 있는 '
                               f'부채라 {v:,.2f} 로 뒤집어 넣었습니다')
        쓴것.add(원.lower())
        덮개[f'D{r}'] = round(float(v), 2)
    # 이름이 바뀌어 옮겨 간 계정은 원래 자리를 0 으로 둡니다 (두 번 세지 않도록)
    for _패, 원 in 패_계정바꿔치기.items():
        for r in sorted(bs당):
            if str(bs당[r].get('A', ('값', ''))[1]).strip() == 원 and f'D{r}' in 덮개:
                덮개[f'D{r}'] = 0.0
    for n, v in QB_BS.items():
        if n.lower() not in 쓴것 and abs(v) > 0.005:
            안내['새계정'].append(('BS', n, v))
    잡기BS = 패_시트계산(bs당, 덮개)
    BS당기 = {}
    for r in sorted(bs당):
        과목 = str(bs당[r].get('H', ('값', ''))[1]).strip()
        if 과목:
            BS당기[과목] = BS당기.get(과목, 0.0) + 잡기BS(f'B{r}')

    # ── 2. IS-당기 : 이번 달 칸을 채우고, 계정과목(국문)별로 모읍니다 ──
    is당 = 패_칸읽기(z, 지도['IS-당기'][0], 공유)
    첫줄, 이름줄 = {}, {}
    for r in sorted(is당):
        a = str(is당[r].get('A', ('값', ''))[1]).strip()
        if a:
            첫줄.setdefault(a, r)
            이름줄.setdefault(a, []).append(r)

    def 앞달합(r):
        s = 0.0
        for c in 패_IS월열[:max(보고월 - 1, 0)]:
            칸 = is당.get(r, {}).get(c)
            if 칸 and 칸[0] == '값':
                s += 패_숫자(칸[1]) or 0.0
        return s

    갈라진 = {}
    for 퀵, (이익, 손실) in 패_FX갈라짐.items():
        갈라진[이익] = (퀵, '이익')
        갈라진[손실] = (퀵, '손실')

    이번칸 = 패_IS월열[보고월 - 1]
    덮개2 = {}
    for r in sorted(is당):
        a = str(is당[r].get('A', ('값', ''))[1]).strip()
        칸 = is당.get(r, {}).get(이번칸)
        if not a or (칸 is not None and 칸[0] == '수식'):
            continue
        if a in 갈라진:
            퀵, 쪽 = 갈라진[a]
            이i, 손i = 첫줄.get(패_FX갈라짐[퀵][0]), 첫줄.get(패_FX갈라짐[퀵][1])
            누적 = QB_IS.get(퀵, 소IS.get(퀵.lower(), 0.0))
            앞 = 앞달합(손i) - 앞달합(이i)
            늘 = 누적 - 앞
            v = max(늘, 0.0) if 쪽 == '손실' else max(-늘, 0.0)
            if r == 이i:
                안내['확인'].append(
                    f'IS 「{퀵}」 : 퀵북 누적 {누적:,.2f} − 양식의 1~{보고월 - 1}월 순액 '
                    f'{앞:,.2f} = {보고월}월 {늘:,.2f} → '
                    f'{"이익" if 늘 < 0 else "손실"} 줄에 넣었습니다')
        elif a in QB_IS or a.lower() in 소IS:
            if 첫줄.get(a) != r:
                v = 0.0
                안내['확인'].append(f'IS 「{a}」 는 양식에 {len(이름줄[a])}줄이라 위 줄에 모두 넣었습니다')
            else:
                # 같은 이름이 여러 줄이면 그 줄들의 앞달 금액을 모두 빼야 이중계산이 없습니다
                앞 = sum(앞달합(rr) for rr in 이름줄[a])
                v = QB_IS.get(a, 소IS.get(a.lower(), 0.0)) - 앞
        else:
            v = 0.0
        덮개2[f'{이번칸}{r}'] = round(float(v), 2)
    잡기IS = 패_시트계산(is당, 덮개2)
    IS당기 = {}
    for r in sorted(is당):
        과목 = str(is당[r].get('S', ('값', ''))[1]).strip()
        if 과목:
            IS당기[과목] = IS당기.get(과목, 0.0) + 잡기IS(f'R{r}')

    # ── 3. BS · IS 표 계산 (계정과목으로 금액을 찾아오는 줄만 넣어 주면 나머지는 수식대로) ──
    표값 = {}
    for 시트, 당기맵 in (('BS', BS당기), ('IS', IS당기)):
        칸들 = 패_칸읽기(z, 지도[시트][0], 공유)
        덮개3 = {}
        for r in sorted(칸들):
            국 = str(칸들[r].get('E', ('값', ''))[1]).strip()
            셀 = 칸들[r].get('I')
            if 국 and 셀 is not None and 셀[0] == '수식' and 'SUMIF' in 셀[1]:
                덮개3[f'I{r}'] = round(float(당기맵.get(국, 0.0)), 2)
        잡기 = 패_시트계산(칸들, 덮개3)
        값 = {}
        for r in sorted(칸들):
            for c in ('H', 'I', 'L', 'M'):
                셀 = 칸들[r].get(c)
                있음 = f'{c}{r}' in 덮개3 or (셀 is not None and 셀[0] == '수식')
                if not 있음:
                    continue
                # 글자를 내놓는 수식(머리글 등)은 건드리지 않습니다
                if (셀 is not None and 셀[0] == '수식' and len(셀) > 2
                        and 셀[2] is not None and 패_숫자(셀[2]) is None):
                    continue
                v = 잡기(f'{c}{r}')
                값[f'{c}{r}'] = round(v, 0 if c in ('H', 'M') else 2)
        표값[시트] = 값

    # ── 3-1. 표지의 기간·환율 이름표를 이번 달로 맞춥니다 ──────
    해 = datetime.date.today().year
    m5 = re.search(r'(20\d{2})', str(패_칸읽기(z, 지도['BS'][0], 공유)
                                     .get(2, {}).get('B', ('값', ''))[1]))
    if m5:
        해 = int(m5.group(1))
    다음달 = datetime.date(해 + (보고월 == 12), 보고월 % 12 + 1, 1)
    월말 = 다음달 - datetime.timedelta(days=1)
    글자값 = {
        'BS': {'B2': f'제? (당) 기 : {월말:%Y년 %m월 %d일}',
               'L5': f'{해 % 100}년 {보고월}월말 환율'},
        'IS': {'B2': f'제?(당)기 : {해}.01.01~{월말:%Y.%m.%d}',
               'L5': f"'{해 % 100}년 평균환율(01월~{보고월}월)"},
    }
    안내['확인'].append(
        f'환율칸(BS M5 · IS M5)은 양식에 있던 값 그대로입니다 — '
        f'{보고월}월말 환율과 1~{보고월}월 평균환율은 직접 넣어 주세요')

    # ── 4. 맨 앞에 붙일 영문 BS · IS (전년도 + 원장상 퀵북 시트) ──
    전기BS = 패_전기표(z, 지도, 공유, 'BS-당기', 'E')
    전기IS = 패_전기표(z, 지도, 공유, 'IS(QB)-기초 검증', 'B')
    영문 = []
    for 시트이름, 퀵북시트, 전기표, 전기말 in (
            (패_영문시트[0], BS시트, 전기BS, 'As of Dec 31, {}'),
            (패_영문시트[1], IS시트, 전기IS, 'January - December {}')):
        제목, 줄들 = 패_영문줄(xls, 퀵북시트, 전기표)
        해 = re.search(r'(20\d{2})', 제목[2] or '')
        해 = int(해.group(1)) if 해 else datetime.date.today().year
        영문.append((시트이름, 제목, 줄들, 전기말.format(해 - 1)))
    안내['영문줄수'] = {n: len(줄) for n, _제, 줄, _전 in 영문}

    # ── 5. 엑셀 만들기 (영문 두 장 + 원래 일곱 장, 수식은 값으로) ──
    결과 = 패_추리기(z, 지도, 표값, 보고월, 영문, 글자값)
    안내['BS계정수'] = len(BS당기)
    안내['IS계정수'] = len(IS당기)
    return 결과, 안내


def 패_추리기(z, 지도, 표값, 보고월, 영문=(), 글자값=None):
    """필요한 일곱 장만 남기고, 수식은 계산된 값으로 바꿔 넣습니다.
       영문 = [(시트이름, 제목들, 줄들, 전기 이름표)] — 맨 앞에 새로 만들어 붙입니다."""
    공유 = 패_공유(z)
    남은시트 = [n for n in 패_남길시트 if n in 지도]
    남길부품, 새시트 = set(), {}
    for 이름 in 패_남길시트:
        if 이름 not in 지도:
            continue
        경로 = 지도[이름][0]
        원xml = z.read(경로).decode('utf-8')
        수식표 = {f'{c}{r}': 셀[1]
                  for r, 줄 in 패_칸읽기(z, 경로, 공유).items()
                  for c, 셀 in 줄.items() if 셀[0] == '수식'}
        새시트[경로] = 패_글자칸(
            패_수식빼기(원xml, 공유, 표값.get(이름, {}), 수식표, 남은시트),
            (글자값 or {}).get(이름, {}))
        남길부품.add(경로)
        rels = 경로.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
        if rels in z.namelist():
            남길부품.add(rels)
            for m in re.finditer(r'Target="([^"]+)"', z.read(rels).decode('utf-8')):
                남길부품.add('xl/' + m.group(1).replace('../', '').lstrip('/'))

    빼는것 = set()
    for n, (경로, _r) in 지도.items():
        if n in 패_남길시트:
            continue
        빼는것.add(경로)
        rels = 경로.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
        빼는것.add(rels)
        if rels in z.namelist():
            for m in re.finditer(r'Target="([^"]+)"', z.read(rels).decode('utf-8')):
                빼는것.add('xl/' + m.group(1).replace('../', '').lstrip('/'))
    빼는것 -= 남길부품

    wb = z.read('xl/workbook.xml').decode('utf-8')
    쓸rid = set()

    def _시트줄(m):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        if 패_풀기(str(a.get('name'))) not in 패_남길시트:
            return ''
        쓸rid.add(a.get('r:id'))
        return re.sub(r'\s*state="\w+"', '', m.group(0))

    쓰던시트id = [int(x) for x in re.findall(r'<sheet\b[^>]*sheetId="(\d+)"',
                                             z.read('xl/workbook.xml').decode('utf-8'))]
    wb = re.sub(r'<sheet\b([^>]*?)/>', _시트줄, wb)
    wb = re.sub(r'<definedNames>.*?</definedNames>', '', wb, flags=re.S)
    # 수식이 살아 있으므로 엑셀이 열릴 때 스스로 다시 계산하도록 표시해 둡니다
    wb = re.sub(r'<calcPr[^>]*/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>',
                wb, count=1)
    wb = re.sub(r'\sactiveTab="\d+"', ' activeTab="0"', wb)

    # 시트 차례는 언제나 BS → IS → CE → MC → CF → CF정산서 → 결산조정분개
    def _차례(m):
        줄들 = re.findall(r'<sheet\b[^>]*?/>', m.group(1))

        def 순서(s):
            a = dict(re.findall(r'([\w:]+)="([^"]*)"', s))
            이름 = 패_풀기(str(a.get('name')))
            return 패_남길시트.index(이름) if 이름 in 패_남길시트 else len(패_남길시트)
        return '<sheets>' + ''.join(sorted(줄들, key=순서)) + '</sheets>'

    wb = re.sub(r'<sheets>(.*?)</sheets>', _차례, wb, count=1, flags=re.S)

    rel = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')

    def _관계줄(m):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        t = 'xl/' + a.get('Target', '').replace('../', '').lstrip('/')
        if t.startswith('xl/worksheets/') and a.get('Id') not in 쓸rid:
            return ''
        return '' if t in 빼는것 else m.group(0)

    rel = re.sub(r'<Relationship\b([^>]*?)/>', _관계줄, rel)
    ct = z.read('[Content_Types].xml').decode('utf-8')
    ct = re.sub(r'<Override PartName="([^"]+)"[^>]*/>',
                lambda m: '' if (m.group(1).lstrip('/') in 빼는것
                                 or m.group(1).lstrip('/') == 'xl/calcChain.xml')
                else m.group(0), ct)

    # ── 영문 두 장을 새로 만들어 맨 앞에 끼웁니다 ──────────────
    붙임, 스타일 = {}, None
    if 영문:
        스타일xml = z.read('xl/styles.xml').decode('utf-8')
        스타일xml, 서식 = 패_스타일더하기(스타일xml)
        스타일 = 스타일xml
        있는이름 = set(z.namelist())
        쓰던rid = set(re.findall(r'Id="(rId\d+)"', rel))
        시트id = max(쓰던시트id or [0])
        머리, 꼬리 = [], []
        for k, (시트이름, 제목, 줄들, 전기이름) in enumerate(영문, 1):
            부품 = f'xl/worksheets/sheetEng{k}.xml'
            while 부품 in 있는이름:
                부품 = 부품.replace('.xml', 'x.xml')
            있는이름.add(부품)
            rid = f'rIdEng{k}'
            while rid in 쓰던rid:
                rid += 'x'
            쓰던rid.add(rid)
            시트id += 1
            붙임[부품] = 패_영문시트XML(제목, 줄들, 전기이름, 서식)
            머리.append(f'<sheet name="{패_엑스(시트이름)}" sheetId="{시트id}" '
                        f'r:id="{rid}"/>')
            꼬리.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats'
                        f'.org/officeDocument/2006/relationships/worksheet" '
                        f'Target="worksheets/{부품.split("/")[-1]}"/>')
            ct = ct.replace('</Types>',
                            f'<Override PartName="/{부품}" ContentType="application/'
                            f'vnd.openxmlformats-officedocument.spreadsheetml.'
                            f'worksheet+xml"/></Types>', 1)
        wb = wb.replace('<sheets>', '<sheets>' + ''.join(머리), 1)
        rel = rel.replace('</Relationships>', ''.join(꼬리) + '</Relationships>', 1)

    buf = io.BytesIO()
    out = zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=6)
    for it in z.infolist():
        n = it.filename
        if n in 빼는것 or n == 'xl/calcChain.xml':
            continue
        if n == 'xl/workbook.xml':
            데이터 = wb.encode('utf-8')
        elif n == 'xl/_rels/workbook.xml.rels':
            데이터 = rel.encode('utf-8')
        elif n == '[Content_Types].xml':
            데이터 = ct.encode('utf-8')
        elif n == 'xl/styles.xml' and 스타일 is not None:
            데이터 = 스타일.encode('utf-8')
        elif n in 새시트:
            데이터 = 새시트[n].encode('utf-8')
        else:
            데이터 = z.read(n)
        out.writestr(it, 데이터)
    for 부품, xml in 붙임.items():
        out.writestr(부품, xml.encode('utf-8'))
    out.close()
    return buf.getvalue()


def 패_글자칸(시트xml, 새글):
    """머리글 같은 글자칸을 갈아 끼웁니다 (서식 번호는 그대로 둡니다)."""
    if not 새글:
        return 시트xml
    앞, 나머지 = 시트xml.split('<sheetData>', 1)
    본문, 뒤 = 나머지.rsplit('</sheetData>', 1)

    def _고치기(m):
        칸, 속 = m.group(1), m.group(2)
        if 칸 not in 새글:
            return m.group(0)
        s = re.search(r'\ss="(\d+)"', 속)
        s = f' s="{s.group(1)}"' if s else ''
        return (f'<c r="{칸}"{s} t="inlineStr"><is>'
                f'<t xml:space="preserve">{패_엑스(새글[칸])}</t></is></c>')

    본문 = re.sub(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', _고치기, 본문, flags=re.S)
    return 앞 + '<sheetData>' + 본문 + '</sheetData>' + 뒤


패_시트참조 = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_가-힣.]+))\s*!")


def 패_바깥본다(수식, 남은시트):
    """수식이 「빼 버리는 시트」를 보고 있으면 True (그런 수식만 값으로 바꿉니다)."""
    for a, b in 패_시트참조.findall(수식 or ''):
        이름 = (a or b).strip()
        if 이름 and 이름 not in 남은시트:
            return True
    return False


def 패_수식빼기(시트xml, 공유, 새값, 수식표=None, 남은시트=()):
    """빼는 시트를 보던 수식만 값으로 바꾸고, 시트 안에서 도는 수식은 살려 둡니다.

      · 수식표 = 공유수식까지 풀어 놓은 {칸: 수식}
      · 새값   = 원장으로 다시 계산한 금액 (수식이 살아남으면 「계산해 둔 값」으로 넣습니다)
    """
    수식표 = 수식표 or {}
    남은 = set(남은시트)
    앞, 나머지 = 시트xml.split('<sheetData>', 1)
    본문, 뒤 = 나머지.rsplit('</sheetData>', 1)

    def _고치기(m):
        칸, 속, 몸 = m.group(1), m.group(2), m.group(3) or ''
        if '<f' not in 몸:
            return m.group(0)
        s = re.search(r'\ss="(\d+)"', 속)
        s = f' s="{s.group(1)}"' if s else ''
        t = re.search(r'\st="(\w+)"', 속)
        t = t.group(1) if t else ''
        vm = re.search(r'<v>(.*?)</v>', 몸, re.S)
        수식 = (수식표.get(칸) or '').strip()
        # TRUE/FALSE·글자를 내놓는 수식은 새로 계산한 숫자를 덮어쓰지 않습니다
        새 = 새값.get(칸) if (칸 in 새값 and t not in ('s', 'str', 'e', 'b')) else None

        if 수식 and not 패_바깥본다(수식, 남은):
            식 = f'<f>{패_엑스(수식.lstrip("="))}</f>'
            if 새 is not None:
                return f'<c r="{칸}"{s}>{식}<v>{float(새)!r}</v></c>'
            if vm is None:
                return f'<c r="{칸}"{s}>{식}</c>'
            꼬리 = f' t="{t}"' if t else ''
            return f'<c r="{칸}"{s}{꼬리}>{식}<v>{vm.group(1)}</v></c>'

        if 새 is not None:
            return f'<c r="{칸}"{s}><v>{float(새)!r}</v></c>'
        if vm is None:
            return f'<c r="{칸}"{s}/>'
        v = vm.group(1)
        if t == 's' and v.isdigit() and int(v) < len(공유):
            v = 공유[int(v)]
        elif t in ('str',):
            v = 패_풀기(v)
        elif t == 'e':
            return f'<c r="{칸}"{s} t="e"><v>{패_엑스(패_풀기(v))}</v></c>'
        elif t == 'b':
            return f'<c r="{칸}"{s} t="b"><v>{v}</v></c>'
        else:
            return f'<c r="{칸}"{s}><v>{v}</v></c>'
        return (f'<c r="{칸}"{s} t="inlineStr">'
                f'<is><t xml:space="preserve">{패_엑스(v)}</t></is></c>')

    본문 = re.sub(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', _고치기, 본문, flags=re.S)
    return 앞 + '<sheetData>' + 본문 + '</sheetData>' + 뒤



# ── 연결패키지 파일 이름 · 버전 (실적 월마다 v1 부터) ─────────────
패키지패턴 = re.compile(r'연결패키지.*?v\s*(\d+)\s*\.xlsx$', re.I)


def 패키지파일이름(연도, 월, 버전):
    return (f'{연도}. {(월 - 1) // 3 + 1}Q_연결패키지_CTKonnect_'
            f'{월}월말_v{버전}.xlsx')


def 패키지목록():
    """자료실에 쌓인 연결패키지 엑셀을 새 것부터 돌려줍니다."""
    try:
        이름들 = [n for n in os.listdir(자료실) if n.lower().endswith('.xlsx')
                  and not n.startswith('~$') and 패키지패턴.search(n)]
    except OSError:
        return []
    목록 = []
    for n in 이름들:
        경로 = os.path.join(자료실, n)
        try:
            st_ = os.stat(경로)
        except OSError:
            continue
        m = re.search(r'(\d{1,2})\s*월말', n)
        목록.append({'이름': n, '경로': 경로,
                     '버전': int(패키지패턴.search(n).group(1)),
                     '월': int(m.group(1)) if m else None,
                     '시각': st_.st_mtime, '크기': st_.st_size})
    목록.sort(key=lambda d: (d['시각'], d['버전']), reverse=True)
    return 목록


def 패키지다음버전(월):
    """같은 실적 월 안에서만 번호를 이어 갑니다 — 달이 바뀌면 저절로 v1 부터입니다."""
    번호 = [d['버전'] for d in 패키지목록() if d['월'] == 월]
    return (max(번호) + 1) if 번호 else 1


@st.cache_data(max_entries=1, show_spinner='연결패키지를 만드는 중입니다...')
def 연결패키지캐시(원장바이트, 양식바이트):
    return 패_연결패키지만들기(원장바이트, 양식바이트)



양식보관 = '최근패키지양식.xlsx'      # 본사에서 받은 연결패키지 원본
패키지보관 = '최근패키지.xlsx'        # 마지막으로 만든 연결패키지 재무제표


def 연결패키지양식칸():
    """재무제표 화면 가운데 — 연결패키지 양식만 올립니다 (원장은 올려 둔 것을 씁니다)."""
    st.markdown('**연결패키지 양식** (지난번 본사 패키지)')
    올림 = st.file_uploader('연결패키지 양식', type=['xlsx'], key='pkg_tpl',
                            label_visibility='collapsed',
                            help='BS-당기 · IS-당기 시트가 든 본사 원본 패키지를 올려 주세요')
    자료남기기(양식보관, 올림)
    양식있음, 양식이름, _크기, _때 = 보관자료(양식보관)
    원장있음, 원장이름, _크기2, _때2 = 보관자료(원장보관)
    if 양식있음 and 올림 is None:
        st.caption(f'양식 : 지난번에 올린 **{양식이름}** 을 그대로 씁니다.')
        바이트 = 보관바이트(양식보관)
        if 바이트:
            st.download_button('⤓  올려 둔 양식 내려받기', data=바이트, file_name=양식이름,
                               width='stretch', key='양식받기',
                               mime=('application/vnd.openxmlformats-officedocument'
                                     '.spreadsheetml.sheet'))
    if 원장있음:
        st.caption(f'원장 : 「실적보고 엑셀작성」 의 **{원장이름}** 을 그대로 씁니다.')
    else:
        st.caption('⚠ 원장이 아직 없습니다 — 「실적보고 엑셀작성」 에서 ① 원장 원본을 '
                   '한 번 올려 주세요.')


def 연결패키지칸():
    """재무제표 화면 오른쪽 — 연결패키지 엑셀 내려받기 · 만들기."""
    보관 = 패키지목록()
    최신 = 보관[0] if 보관 else None
    st.caption('연결패키지 재무제표')
    if 최신 is not None:
        try:
            with open(최신['경로'], 'rb') as f:
                st.download_button(
                    f"⤓  {최신['월'] or ''}월말 v{최신['버전']} 내려받기", data=f.read(),
                    file_name=최신['이름'], width='stretch', type='primary',
                    mime=('application/vnd.openxmlformats-officedocument'
                          '.spreadsheetml.sheet'),
                    help='BS(영문) · IS(영문) · BS · IS · CE · MC · CF · '
                         'CF정산서 · 결산조정분개 아홉 장')
        except OSError:
            st.button('⤓  내려받기', width='stretch', disabled=True)
    elif 보관자료(패키지보관)[0]:
        # 자료실이 비어 있어도(웹사이트가 다시 켜진 뒤) 마지막으로 만든 것은 드립니다
        _있, 팩이름, 팩크기, 팩때 = 보관자료(패키지보관)
        바이트 = 보관바이트(패키지보관)
        if 바이트:
            st.download_button(
                '⤓  마지막 패키지 내려받기', data=바이트, file_name=팩이름,
                width='stretch', type='primary', key='마지막패키지받기',
                mime=('application/vnd.openxmlformats-officedocument'
                      '.spreadsheetml.sheet'),
                help=f'{팩이름}  ·  {팩크기/1e6:,.1f}MB  ·  '
                     f'{datetime.datetime.fromtimestamp(팩때):%Y.%m.%d %H:%M} 에 만듦')
        else:
            st.button('⤓  내려받기', width='stretch', disabled=True)
    else:
        st.button('⤓  내려받기', width='stretch', disabled=True,
                  help='아직 만든 패키지가 없습니다. 가운데에 양식을 올리고 만들어 주세요')

    원장있음 = 보관자료(원장보관)[0]
    양식있음 = 보관자료(양식보관)[0]
    준비 = 원장있음 and 양식있음

    def _누름():
        st.session_state['패키지생성요청'] = True

    st.button('⚙  연결패키지 재무제표 만들기', width='stretch', disabled=not 준비,
              on_click=_누름, key='pkg_make', type='primary' if 준비 else 'secondary',
              help=('가운데에 연결패키지 양식을 올리면 켜집니다' if 원장있음
                    else '「실적보고 엑셀작성」 에서 원장을 먼저 올려 주세요'))
    if st.session_state.pop('패키지생성요청', False) and 준비:
        try:
            바이트, 안내 = 연결패키지캐시(보관바이트(원장보관), 보관바이트(양식보관))
            월 = 안내['보고월']
            버전 = 패키지다음버전(월)
            이름 = 패키지파일이름(당해연도, 월, 버전)
            저장 = ''
            try:
                os.makedirs(자료실, exist_ok=True)
                with open(os.path.join(자료실, 이름), 'wb') as f:
                    f.write(바이트)
                저장 = os.path.abspath(os.path.join(자료실, 이름))
            except OSError as e:
                st.caption(f'※ 자료실에 저장하지 못했습니다({e})')
            _자리에쓰기(패키지보관, 바이트, 이름)      # 언제든 다시 받을 수 있게
            st.session_state['패키지결과'] = (이름, 바이트, 안내, 저장)
        except Exception as e:
            st.session_state['패키지오류'] = str(e)

    오류 = st.session_state.pop('패키지오류', None)
    if 오류:
        st.caption(f'⚠ 만들지 못했습니다 — {오류}')
    결과 = st.session_state.get('패키지결과')
    if 결과:
        이름, 바이트, 안내, 저장 = 결과
        st.download_button(f'⤓  {이름[:28]}… 내려받기', data=바이트, file_name=이름,
                           width='stretch', key='pkg_dl_new',
                           mime=('application/vnd.openxmlformats-officedocument'
                                 '.spreadsheetml.sheet'))
        st.caption(f"{안내['보고월']}월말 기준 · 계정과목 BS {안내['BS계정수']}개 · "
                   f"IS {안내['IS계정수']}개" + (f' · 저장 {저장}' if 저장 else ''))
        for 말 in 안내.get('확인', [])[:8]:
            st.caption(f'※ {말}')


def 음수표기(글자):
    """마이너스 금액은 회계 방식대로 (괄호) 안에 빨간 글씨로 보여줍니다."""
    s = str(글자)
    if s[:1] in ('-', '−'):
        return f'<span class="neg">({s[1:]})</span>'
    return s


금액나눔 = 1          # 씨티케이커넥트 화면은 1,000 (천원) — 아래 표시나눔 줄에서 바꿈 (2026-11-03)


def 금액(v, 자릿수=0, 괄호=True):
    """괄호=False 면 캡션처럼 HTML 을 못 쓰는 곳에서 쓸 수 있게 그냥 글자만 돌려줍니다."""
    if v is None or pd.isna(v):
        return '-'
    v = v / 금액나눔
    if abs(v) < 0.5 / (10 ** 자릿수):        # 반올림하면 0 이 되는 금액은 '-' 로
        return '-'
    t = f'{v:,.{자릿수}f}'
    return 음수표기(t) if 괄호 else t


전년없음 = False


def 증감HTML(현재, 이전, 라벨='전월대비'):
    if 이전 is None or pd.isna(이전) or 이전 == 0:
        if 전년없음 and '전년' in str(라벨) or (전년없음 and str(라벨).startswith('20')):
            return ''        # CTKonnect : 1기 법인이라 전년 비교 칸은 비워 둡니다
        return '<span class="muted">- 비교불가</span>'
    diff = 현재 - 이전
    pct = diff / abs(이전) * 100
    화살표 = '▲' if diff >= 0 else '▼'
    색 = 'up' if diff >= 0 else 'down'
    return f'<span class="{색}">{화살표} {라벨} {abs(pct):,.1f}%</span>'


def 한줄설명(행들):
    """그 거래처에서 무슨 일이 있었는지 20자 안팎으로.

    원장의 한글 분류(활동세부 → 대분류)만 씁니다. 영문 적요를 그대로 내보내면
    브라우저 번역 기능이 엉뚱한 말로 바꿔 버려서, 한글 분류가 없으면 비워 둡니다.
    """
    for 열 in ('활동세부', '활동분류(대분류)'):
        if 열 in 행들.columns:
            값 = 행들[열].dropna().astype(str).str.strip()
            값 = 값[값.ne('') & 값.ne('nan') & 값.ne('0')]
            값 = 값[값.apply(lambda t: any('가' <= c <= '힣' for c in t))]
            if len(값):
                t = 값.value_counts().index[0]
                return t if len(t) <= 22 else t[:20] + '…'
    return ''


def 세부패널(대상, 값컬럼, 부호=1, 상위N=7):
    """항목 하나를 골랐을 때 거래처별 누적 내역을 보여주는 카드 내용(표 + 거래 내용)을 만듭니다.
    매출이든 비용이든 금액이 큰 상위 N곳만 적고 나머지는 「기타」로 한 줄에 모읍니다.
    당월 실적집계·대시보드 화면에서 공통으로 씁니다."""
    if 대상 is None or 대상.empty:
        return '<div class="muted">해당 기간에 거래가 없습니다.</div>'
    집계 = (대상.groupby(대상['Name'].fillna('(거래처 미기재)'))[값컬럼].sum() * 부호)
    집계 = 집계[집계.abs() > 0.5].sort_values(ascending=False)
    if 집계.empty:
        return '<div class="muted">해당 기간에 거래가 없습니다.</div>'
    앞 = 집계.iloc[:상위N]
    기타합 = float(집계.iloc[상위N:].sum())
    총액 = float(집계.sum()) or 1.0
    줄 = ''
    for 이름2, v in 앞.items():
        줄 += (f'<tr><td class="lft notranslate" translate="no">{이름2}</td>'
               f'<td>{금액(v)}</td><td class="muted">{v / 총액 * 100:,.1f}%</td></tr>')
    if abs(기타합) > 0.5:
        줄 += (f'<tr><td class="lft">기타 ({len(집계) - 상위N}곳 합계)</td>'
               f'<td>{금액(기타합)}</td><td class="muted">{기타합 / 총액 * 100:,.1f}%</td></tr>')
    줄 += (f'<tr style="font-weight:700"><td class="lft">합계</td>'
           f'<td>{금액(float(집계.sum()))}</td><td>100.0%</td></tr>')
    # ── 거래 내용 : 거래처마다 「무슨 거래인지」 + 월별 금액 ─────────────
    #    어느 달에 늘었는지 가로로 훑어볼 수 있게, 달을 열로 놓고 옆으로 밀리는 상자에 담습니다.
    이름칸 = 대상['Name'].fillna('(거래처 미기재)')
    월들 = []
    if '월' in 대상.columns:
        월들 = sorted({int(m) for m in pd.to_numeric(대상['월'], errors='coerce').dropna()})

    def 칸값(v):
        return 금액(v) if abs(v) > 0.5 else '<span class="muted">-</span>'

    if 월들:
        피벗 = (대상.assign(_이름=이름칸)
                .pivot_table(index='_이름', columns='월', values=값컬럼,
                             aggfunc='sum', fill_value=0.0)
                .reindex(columns=월들, fill_value=0.0) * 부호)
        머리 = ''.join(f'<th>{m}월</th>' for m in 월들)
        본문 = ''
        for 이름2 in 앞.index:
            묶음 = 대상[이름칸.eq(이름2)]
            구분 = 한줄설명(묶음)
            달값 = [float(피벗.loc[이름2, m]) if 이름2 in 피벗.index else 0.0 for m in 월들]
            본문 += (f'<tr><td class="lft"><b class="notranslate" translate="no">{이름2}</b>'
                     + (f'<span class="cap">{구분}</span>' if 구분 else '')
                     + '</td>'
                     + ''.join(f'<td>{칸값(v)}</td>' for v in 달값)
                     + f'<td class="sum">{금액(float(앞[이름2]))}</td></tr>')
        나머지 = [i for i in 피벗.index if i not in set(앞.index)]
        if 나머지 and abs(기타합) > 0.5:
            달값 = [float(피벗.loc[나머지, m].sum()) for m in 월들]
            본문 += (f'<tr><td class="lft"><b>기타 ({len(나머지)}곳 합계)</b></td>'
                     + ''.join(f'<td>{칸값(v)}</td>' for v in 달값)
                     + f'<td class="sum">{금액(기타합)}</td></tr>')
        달합 = [float(피벗[m].sum()) for m in 월들]
        본문 += ('<tr class="tot"><td class="lft"><b>합계</b></td>'
                 + ''.join(f'<td>{칸값(v)}</td>' for v in 달합)
                 + f'<td class="sum">{금액(float(집계.sum()))}</td></tr>')
        설명 = (f'<div class="mscroll"><table class="mtab">'
                f'<thead><tr><th class="lft">거래처</th>{머리}<th class="sum">누적</th></tr></thead>'
                f'<tbody>{본문}</tbody></table></div>')
    else:
        설명 = ''
        for 이름2 in 앞.index:
            묶음 = 대상[이름칸.eq(이름2)]
            설명 += (f'<div class="memo"><b class="notranslate" translate="no">{이름2}</b>'
                     f'<span>{한줄설명(묶음)}</span></div>')

    # 화면이 좁아도 카드 밖으로 잘려 나가지 않도록 스크롤 상자에 담습니다
    return (f'<div class="scrollx"><table class="detail">'
            f'<colgroup><col style="width:47%"><col style="width:32%"><col style="width:21%"></colgroup>'
            f'<thead><tr><th class="lft">거래처</th><th>누적 금액</th>'
            f'<th>비중</th></tr></thead><tbody>{줄}</tbody></table></div>'
            f'<div class="sub" style="margin:14px 0 6px">거래 내용 '
            f'<span class="calc">(달마다 얼마인지 · 옆으로 밀어 보세요)</span></div>{설명}')


# ══════════════════════════════════════════════════════════════
# 2. 사이드바
# ══════════════════════════════════════════════════════════════
if 'menu' not in st.session_state:
    st.session_state['menu'] = 'CTKonnect'

st.sidebar.markdown("""
<div class="brand-box">
    <div class="brand-title">CTKonnect 실적보고</div>
    <div class="brand-sub">GROUP FINANCE VIEW</div>
</div>
""", unsafe_allow_html=True)

# 제목 바로 아래에 「최종자료 ○년 ○월」 을 띄울 자리 — 자료를 읽은 뒤 채웁니다
자료월칸 = st.sidebar.empty()

메뉴목록 = ['실적보고', '미수채권 관리', '계정과목 상세', '데이터 점검',
            '연결재무제표 패키지', '실적보고 엑셀작성', '월간실적']
# 「실적보고」를 누르면 아래 세 화면이 펼쳐집니다 (맨 처음은 누적 실적보고)
실적하위 = ['CTKonnect', '씨티케이커넥트', '연결 손익계산서']   # 2026-10-29 커넥트는 미국 · 한국 두 법인        # 당월·월별 실적보고는 누적 화면의 손익계산서(연도·판관비 구분 선택)로 합침 (2026-10-27)
# 「연결재무제표 패키지」도 같은 방식으로 네 화면을 거느립니다
패키지하위 = ['재무제표', '주석사항']
# 재무제표 화면 안에서 고르는 세 가지 (자동 번역을 막으려고 영문 약칭을 앞에 답니다)
재무제표탭 = ['BS  재무상태표', 'IS  손익계산서', 'MC  제조원가명세서']
# 연결패키지 BS·IS 시트에 국문·영문 두 가지 이름이 나란히 있어, 화면도 같게 나눕니다
표기구분 = ['국문 (한글 계정과목)', 'English (account names)']

# ── 연결패키지 BS 서식 : (번호, 국문, 영문, 들여쓰기, 더할줄) ──
#    더할줄이 None 이면 계정과목 이름으로 금액을 찾아오는 줄,
#    아니면 [(줄번호, 부호)] 대로 위 줄들을 더하고 빼는 줄입니다.
BS서식 = [
    ('', '자산', '', 3, None),
    ('Ⅰ.', '유동자산', 'Current assets', 0, [(2, 1), (21, 1)]),
    ('(1)', '당좌자산', 'Quick Assets', 1, [(3, 1), (4, 1), (5, 1), (6, 1), (7, 1), (8, 1), (9, 1), (10, 1), (11, 1), (12, 1), (13, 1), (14, 1), (15, 1), (16, 1), (17, 1), (18, 1), (19, 1), (20, 1)]),
    ('1.', '현금및현금성자산', 'Cash and cash equivalents', 2, None),
    ('2.', '단기금융예치금', 'Short-term Financial Instruments', 2, None),
    ('3.', '기타금융자산(유동)', 'Current fair value financial asset', 2, None),
    ('4.', '매출채권', 'Trade Receivable, gross', 2, None),
    ('', '매출채권-대손충당금', 'Allowance for Doubtful Accounts', 3, None),
    ('5.', '단기대여금', 'Short-term loans', 2, None),
    ('', '단기대여금-대손충당금', 'Allowance for Doubtful Accounts', 3, None),
    ('6.', '미수금', 'Other Receivables, gross', 2, None),
    ('', '미수금-대손충당금', 'Allowance for Doubtful Accounts', 3, None),
    ('7.', '미수수익', 'Accrued Income, gross', 2, None),
    ('', '미수수익-대손충당금', 'Allowance for Doubtful Accounts', 3, None),
    ('8.', '선급금', 'Advance payments, gross', 2, None),
    ('9.', '선급비용', 'Prepaid Expense', 2, None),
    ('10.', '당기법인세자산', 'Current tax assets', 2, None),
    ('11.', '부가세대급금', 'Prepaid Value Added Tax', 2, None),
    ('12.', '파생상품금융자산', 'Current derivative financial assets', 2, None),
    ('13.', '단기보증금(유동)', 'Short-term deposits Provided', 2, None),
    ('14.', '매각예정자산', 'disposal groups classified as held for sale', 2, None),
    ('(2)', '재고자산', 'Inventory', 1, [(22, 1), (23, 1), (24, 1), (25, 1), (26, 1), (27, 1), (28, 1), (29, 1), (30, 1)]),
    ('1.', '상품', 'Merchandise, gross', 2, None),
    ('', '평가충당금(상품)', 'Accumulated Impairment loss', 3, None),
    ('2.', '제품', 'Finished goods, gross', 2, None),
    ('', '평가충당금(제품)', 'Accumulated Impairment loss', 3, None),
    ('3.', '재공품', 'Work in progress, gross', 2, None),
    ('', '평가충당금(재공품)', 'Accumulated Impairment loss', 3, None),
    ('4.', '원재료', 'Raw materials, gross', 2, None),
    ('', '평가충당금(원재료)', 'Accumulated Impairment loss', 3, None),
    ('5.', '미착품', 'Goods in transit, gross', 2, None),
    ('Ⅱ.', '비유동자산', 'Non-current assets', 0, [(32, 1), (38, 1), (58, 1), (62, 1), (70, 1), (73, 1), (75, 1)]),
    ('(1)', '투자자산', 'Investments', 1, [(33, 1), (34, 1), (35, 1), (36, 1), (37, 1)]),
    ('1.', '장기금융예치금', 'Llong-term financial instruments', 2, None),
    ('2.', '기타금융자산(비유동)', 'Non-current fair value financial asset', 2, None),
    ('3.', '관계기업투자자산', 'Investments in associates', 2, None),
    ('4.', '종속기업투자자산', 'Investments in subsidiaries', 2, None),
    ('5.', '장기대여금', 'Long-term loans', 2, None),
    ('(2)', '유형자산', 'Property, plant and equipment', 1, [(39, 1), (40, 1), (41, 1), (42, 1), (43, 1), (44, 1), (45, 1), (46, 1), (47, 1), (48, 1), (49, 1), (50, 1), (51, 1), (52, 1), (53, 1), (54, 1), (55, 1), (56, 1), (57, 1)]),
    ('1.', '토지', 'Land, gross', 2, None),
    ('2.', '건물', 'Buildings, gross', 2, None),
    ('', '건물-감가상각누계액', 'Accumulated depreciation', 3, None),
    ('3.', '시설장치', 'facilities', 2, None),
    ('', '시설장치-감가상각누계액', 'Accumulated depreciation', 3, None),
    ('4.', '기계장치', 'Machinery', 2, None),
    ('', '기계장치-감가상각누계액', 'Accumulated depreciation', 3, None),
    ('5.', '차량운반구', 'Vehicles, gross', 2, None),
    ('', '차량운반구-감가상각누계액', 'Accumulated depreciation', 3, None),
    ('6.', '공구와기구', 'Tools, gross', 2, None),
    ('', '공구와기구-감가상각누계액', 'Accumulated depreciation', 3, None),
    ('7.', '비품', 'Office equipment, gross', 2, None),
    ('', '비품-감가상각누계액', 'Accumulated depreciation', 3, None),
    ('8.', '금형', 'Molds, gross', 2, None),
    ('', '금형-감가상각누계액', 'Accumulated depreciation', 3, None),
    ('9.', '건설중인자산', 'Construction in progress, gross', 2, None),
    ('', '정부보조금-건설중인자산', 'Government grants_Construction in progress, gross', 3, None),
    ('10.', '정부보조금-기계장치', 'Government grants_Machinery', 2, None),
    ('', '정부보조금-기계장치-감가상각누계액', 'Accumulated depreciation', 3, None),
    ('(3)', '투자부동산', 'Investment property', 1, [(59, 1), (60, 1), (61, 1)]),
    ('1.', '투자부동산-토지', 'Investment property - land, gross', 2, None),
    ('2.', '투자부동산-건물', 'Investment property - buildings, gross', 2, None),
    ('', '투자부동산-건물감가상각누계액', 'Accumulated depreciation', 3, None),
    ('(4)', '무형자산', 'Intangible assets', 1, [(63, 1), (64, 1), (65, 1), (66, 1), (67, 1), (68, 1), (69, 1)]),
    ('1.', '영업권', 'Goodwill', 2, None),
    ('2.', '특허권', 'Patents', 2, None),
    ('3.', '상표권', 'Trademarks', 2, None),
    ('4.', '의장권', 'Patents-Design', 2, None),
    ('5.', '소프트웨어', 'Software', 2, None),
    ('6.', '기타의무형자산', 'Other intangible assets', 2, None),
    ('7.', '시설이용권', 'Facility usage rights', 2, None),
    ('(5)', '사용권자산', 'Lease assets', 1, [(71, 1), (72, 1)]),
    ('', '사용권자산(비유동)', 'Non-currment lease assets', 3, None),
    ('', '사용권자산-상각누계액', 'Accumulated depreciation', 3, None),
    ('(6)', '순확정급여자산', 'Defined benefit assets', 1, [(74, 1)]),
    ('', '사외적립자산(자산)', 'Fair value of plan assets', 3, None),
    ('(7)', '기타비유동자산', 'Other non-current assets', 1, [(76, 1), (77, 1), (78, 1), (79, 1), (80, 1), (81, 1), (82, 1), (83, 1), (84, 1)]),
    ('1.', '장기보증금', 'Long-term deposits provided', 2, None),
    ('', '보증금현재가치할인차금', 'Present Value Discount', 3, None),
    ('', '장기보증금-대손충당금', 'Allowance for Doubtful Accounts', 3, None),
    ('2.', '장기선급금', 'Advance payments, gross', 2, None),
    ('3.', '장기선급비용', 'Prepaid Expense', 2, None),
    ('4.', '파생상품금융자산(비유동)', 'Non-currment derivative financial assets', 2, None),
    ('5.', '이연법인세자산(비유동)', 'Non-current deferred tax assets', 2, None),
    ('6.', '장기미수수익', 'Long-term accrued Income, gross', 2, None),
    ('', '장기미수수익-대손충당금', 'Allowance for Doubtful Accounts', 3, None),
    ('', '자산총계', 'Total Assets', 3, [(1, 1), (31, 1)]),
    ('', '부채', '', 3, None),
    ('Ⅰ.', '유동부채', 'Current liabilities', 0, [(88, 1), (89, 1), (90, 1), (91, 1), (92, 1), (93, 1), (94, 1), (95, 1), (96, 1), (97, 1), (98, 1), (99, 1), (100, 1), (101, 1), (102, 1), (103, 1)]),
    ('1.', '매입채무', 'Trade Payables, gross', 2, None),
    ('2.', '단기차입금', 'Short term borrowings', 2, None),
    ('3.', '미지급금', 'Other Payables', 2, None),
    ('4.', '미지급비용', 'Accrued expenses', 2, None),
    ('5.', '리스부채(유동)', 'Current finance lease liabilities', 2, None),
    ('6.', '예수금', 'Withholdings banks', 2, None),
    ('7.', '부가세예수금', 'Value Added Tax Withheld', 2, None),
    ('8.', '선수금', 'Advances Customers', 2, None),
    ('9.', '금융보증부채(유동)', 'Current financial guarantee liabilities', 2, None),
    ('10.', '유동성장기부채', 'Current Portion of Long-term Liabilities', 2, None),
    ('11.', '미지급배당금', 'Dividends Payable', 2, None),
    ('12.', '당기법인세부채', 'Current tax liabilities', 2, None),
    ('13.', '당기손익인식금융부채', 'Fair value financial liabilities', 2, None),
    ('14.', '파생상품금융부채(유동)', 'Currment derivative financial liabilities', 2, None),
    ('15.', '전환사채(유동)', '', 2, None),
    ('16.', '하자보수충당부채', 'Provision for Construction Warranties', 2, None),
    ('Ⅱ.', '비유동부채', 'Non-current liabilities', 0, [(105, 1), (106, 1), (107, 1), (108, 1), (109, 1), (110, 1), (111, 1), (112, 1), (113, 1), (114, 1), (115, 1)]),
    ('1.', '장기차입금', 'Long-term borrowings, gross', 2, None),
    ('2.', '임대보증금', 'Guarantee Deposit Rent, gross', 2, None),
    ('3.', '리스부채(비유동)', 'Non-current finance lease liabilities', 2, None),
    ('4.', '확정급여채무', 'Defined benefit obligation', 2, None),
    ('', '사외적립자산', 'Fair value of plan assets', 3, None),
    ('5.', '이연법인세부채', 'Deferred tax liabilities', 2, None),
    ('6.', '장기미지급금', 'Long-term other payables', 2, None),
    ('7.', '기타장기급여채무', 'Long-term benefit obligation', 2, None),
    ('8.', '금융보증부채(비유동)', 'Non-current financial guarantee liabilities', 2, None),
    ('9.', '파생상품금융부채(비유동)', 'Non-currment derivative financial liabilities', 2, None),
    ('10.', '전환사채(비유동)', 'Convertible bonds', 2, None),
    ('', '부채총계', 'Total liabilities', 3, [(87, 1), (104, 1)]),
    ('', '자본', '', 3, None),
    ('Ⅰ.', '자본금', 'Issued capital', 0, [(119, 1), (120, 1)]),
    ('1.', '보통주자본금', 'Issued capital of common stock', 2, None),
    ('2.', '우선주자본금', 'Issued capital of preferred stock', 2, None),
    ('Ⅱ.', '기타불입자본', 'Capital surplus', 0, [(122, 1), (123, 1), (124, 1)]),
    ('1.', '주식발행초과금', 'Share premium', 2, None),
    ('2.', '자기주식처분이익', 'Gains on Sale of Treasury Stock', 2, None),
    ('3.', '기타자본잉여금', 'Other capital surplus', 2, None),
    ('Ⅲ.', '자본조정', 'Capital Adjustments', 0, [(126, 1), (127, 1), (128, 1), (129, 1), (130, 1), (131, 1)]),
    ('1.', '주식할인발행차금', 'Discount on Stock Issuance', 2, None),
    ('2.', '기타자본조정', 'Other capital adjustments', 2, None),
    ('3.', '자기주식', 'Treasury Stock', 2, None),
    ('4.', '주식매수선택권', 'Stock Option', 2, None),
    ('5.', '자기주식처분이익', 'Gains sale treasury stock', 2, None),
    ('6.', '자기주식처분손실', 'Losses sale treasury stock', 2, None),
    ('IV.', '기타포괄손익누계액', 'Other Comprehensive income/loss accumulated amount', 0, [(133, 1), (134, 1), (135, 1), (136, 1), (137, 1), (138, 1)]),
    ('1.', '매도가능증권평가이익', 'Gains on Valuation of available-for-sale financial assets', 2, None),
    ('2.', '매도가능증권평가손실', 'Losses on Valuation of available-for-sale financial assets', 2, None),
    ('3.', '해외사업환산이익', 'Cumulative gains of Foreign operation', 2, None),
    ('4.', '해외사업환산손실', 'Cumulative losses of Foreign operation', 2, None),
    ('5.', '지분법자본변동', 'Gains on financial assets as measured at fair value', 2, None),
    ('6.', '부의지분법자본변동', 'Losses on financial assets as measured at fair value', 2, None),
    ('IV.', '이익잉여금', 'Retained earnings', 0, [(140, 1), (141, 1), (142, 1)]),
    ('1.', '이익준비금', 'Legal Reserve', 2, None),
    ('2.', '미처분이익잉여금', 'Unappropriated Retained Earnings', 2, None),
    ('3.', '확정급여부채재측정요소', 'Reclassification factor on defined benefit plans', 2, None),
    ('IV.', '비지배지분', 'Non-controlling interests', 0, None),
    ('', '자본총계', 'Total Equities', 3, [(118, 1), (121, 1), (125, 1), (132, 1), (139, 1), (143, 1)]),
    ('', '부채및자본총계', 'Total equity and liabilities', 3, [(144, 1), (116, 1)]),
]

# BS 서식에서 「부채」 머리줄 아래에 있는 계정과목 (대변이 늘어나는 쪽)
BS부채자본 = set()
_부채구역 = False
for _번, _국, _영, _수준, _더할 in BS서식:
    if not _영 and _국 in ('부채', '자본'):
        _부채구역 = True
    elif not _영 and _국 == '자산':
        _부채구역 = False
    if _부채구역 and _국:
        BS부채자본.add(_국)

# ── 연결패키지 IS 서식 : (번호, 국문, 영문, 들여쓰기, 더할줄) ──
#    더할줄이 None 이면 계정과목 이름으로 금액을 찾아오는 줄,
#    아니면 [(줄번호, 부호)] 대로 위 줄들을 더하고 빼는 줄입니다.
IS서식 = [
    ('Ⅰ.', '매 출 액', 'Sales', 0, [(1, 1), (2, 1), (3, 1), (4, 1), (5, 1)]),
    ('1.', '상품매출', 'Sales of Merchandise', 2, None),
    ('2.', '제품매출', 'Sales of Finished Goods', 2, None),
    ('3.', '패키지매출', 'Sales of Package', 2, None),
    ('4.', '용역매출', 'Revenues-Services', 2, None),
    ('5.', '기타매출', 'Sales of Other', 2, None),
    ('Ⅱ.', '매 출 원 가', 'Cost of sales', 0, [(7, 1), (8, 1), (9, 1), (10, 1), (11, 1), (12, 1), (13, 1), (14, 1), (15, 1), (16, 1), (17, 1), (18, 1)]),
    ('1.', '상품매출원가', 'Cost of Merchandise Sold', 2, None),
    ('', '기초상품재고액', '', 3, None),
    ('', '당기상품입고액', '', 3, None),
    ('', '기말상품재고액', '', 3, None),
    ('2.', '제품매출원가', 'Cost of Finished Goods Sold', 2, None),
    ('', '기초제품재고액', '', 3, None),
    ('', '당기제품제조원가', '', 3, None),
    ('', '기말제품재고액', '', 3, None),
    ('', '타계정출고액', '', 3, None),
    ('3.', '패키지매출원가', '', 2, None),
    ('4.', '용역매출원가', 'Cost of Service', 2, None),
    ('5.', '기타매출원가', 'Cost of Other', 2, None),
    ('Ⅲ.', '매출총이익', 'Gross Profits', 0, [(0, 1), (6, -1)]),
    ('Ⅳ.', '판매비와관리비', 'Selling and Administrative Expenses', 0, [(21, 1), (22, 1), (23, 1), (24, 1), (25, 1), (26, 1), (27, 1), (28, 1), (29, 1), (30, 1), (31, 1), (32, 1), (33, 1), (34, 1), (35, 1), (36, 1), (37, 1), (38, 1), (39, 1), (40, 1), (41, 1), (42, 1), (43, 1), (44, 1), (45, 1), (46, 1), (47, 1), (48, 1), (49, 1), (50, 1), (51, 1), (52, 1)]),
    ('1.', '급여', 'Salaries wages', 2, None),
    ('2.', '퇴직급여', 'Provision for severance indemnities', 2, None),
    ('3.', '복리후생비', 'Employee benefits', 2, None),
    ('4.', '주식보상비용', 'Stock option expenses', 2, None),
    ('5.', '여비교통비', 'Travel expenses', 2, None),
    ('6.', '접대비', 'Entertainment expenses', 2, None),
    ('7.', '통신비', 'Communication expenses', 2, None),
    ('8.', '수도광열비', 'Utility expenses', 2, None),
    ('9.', '전력비', 'Electricity expenses', 2, None),
    ('10.', '세금과공과', 'Taxes dues', 2, None),
    ('11.', '감가상각비', 'Depreciation expense', 2, None),
    ('12.', '지급임차료', 'Rental expenses', 2, None),
    ('13.', '수선비', 'Repair expenses', 2, None),
    ('14.', '보험료', 'Insurance premiums', 2, None),
    ('15.', '차량유지비', 'Vehicle maintenance expenses', 2, None),
    ('16.', '경상연구개발비', 'Ordinary development expense', 2, None),
    ('17.', '운반비', 'Freight expenses', 2, None),
    ('18.', '교육훈련비', 'Training Expenses', 2, None),
    ('19.', '도서인쇄비', 'Publication expenses', 2, None),
    ('20.', '사무용품비', 'office supply expenses', 2, None),
    ('21.', '소모품비', 'Supply expenses', 2, None),
    ('22.', '지급수수료', 'Commissions & Professional/Outside service Fees', 2, None),
    ('23.', '광고선전비', 'Advertising expenses', 2, None),
    ('24.', '판매촉진비', 'Sales Promotional Expenses', 2, None),
    ('25.', '대손상각비', 'Bad Debt Expenses', 2, None),
    ('26.', '건물관리비', 'Maintenance expenses on buildings', 2, None),
    ('27.', '수출제비용', 'Export expenses', 2, None),
    ('28.', '무형자산상각비', 'Amortisation expense', 2, None),
    ('29.', '견본비', 'Samples Expenses', 2, None),
    ('30.', '해외시장개척비', 'Overseas Marketing Expenses', 2, None),
    ('31.', '하자보수비', 'Warranty Expenses', 2, None),
    ('32.', '잡비', 'Miscellaneous Expenses', 2, None),
    ('Ⅴ.', '영업이익', 'Operating Profits', 0, [(19, 1), (20, -1)]),
    ('Ⅵ.', '기타수익', 'Other Revenues', 0, [(55, 1), (56, 1), (57, 1), (58, 1), (59, 1), (60, 1), (61, 1), (62, 1)]),
    ('1.', '유형자산처분이익', 'Gains on disposals of property, plant and equipment', 2, None),
    ('2.', '사용권자산처분이익', 'Gains on disposals of lease assets', 2, None),
    ('3.', '대손충당금환입', 'Reversal Allowance Doubtful Accounts', 2, None),
    ('4.', '잡이익', 'Miscellaneous Income', 2, None),
    ('5.', '염가매수이익', 'Amortization on negative goodwill', 2, None),
    ('6.', '관계기업투자주식처분이익', 'Gains on sale of investments in associates financial income', 2, None),
    ('7.', '종속기업투자자산처분이익', 'Gains on sale of investments in subsidiaries', 2, None),
    ('8.', '전기오류수정이익', '', 2, None),
    ('Ⅶ.', '기타비용', 'Other Expenses', 0, [(64, 1), (65, 1), (66, 1), (67, 1), (68, 1), (69, 1), (70, 1), (71, 1), (72, 1), (73, 1), (74, 1)]),
    ('1.', '유형자산처분손실', 'Losses on disposals of property, plant and equipment', 2, None),
    ('2.', '유형자산폐기손실', '', 2, None),
    ('3.', '유형자산손상차손', 'Impairment loss of property, plant and equipment', 2, None),
    ('4.', '무형자산처분손실', '', 2, None),
    ('5.', '관계기업투자주식처분손실', '', 2, None),
    ('6.', '종속기업투자자산처분손실', '', 2, None),
    ('7.', '종속기업투자자산손상차손', '', 2, None),
    ('8.', '관계기업투자자산손상차손', '', 2, None),
    ('9.', '기타의대손상각비', 'Other Bad Debt Expenses', 2, None),
    ('10.', '기부금', 'Donations', 2, None),
    ('11.', '잡손실', 'Miscellaneous Losses', 2, None),
    ('Ⅷ.', '금융수익', 'Finance income', 0, [(76, 1), (77, 1), (78, 1), (79, 1), (80, 1), (81, 1), (82, 1), (83, 1), (84, 1), (85, 1)]),
    ('1.', '이자수익', 'Interest income', 2, None),
    ('2.', '배당금수익', 'Dividend Income', 2, None),
    ('3.', '외환차익', 'Gains Foreign Currency Transactions', 2, None),
    ('4.', '외화환산이익', 'Gain on foreign exchagne translations', 2, None),
    ('5.', '보증수익', '', 2, None),
    ('6.', '기타금융자산평가이익', 'Gains on valuation of fair value financial asset', 2, None),
    ('7.', '기타금융자산처분이익', 'Gains on disposal of fair value financial asset', 2, None),
    ('8.', '매도가능증권처분이익', 'Gains on disposal of available-for-sale financial assets', 2, None),
    ('9.', '파생상품평가이익', 'Gains on valuation of derivative financial assets', 2, None),
    ('10.', '파생상품거래이익', 'Gains on disposal of derivative financial assets', 2, None),
    ('Ⅸ.', '금융비용', 'Finance expense', 0, [(87, 1), (88, 1), (89, 1), (90, 1), (91, 1), (92, 1), (93, 1), (94, 1), (95, 1), (96, 1), (97, 1)]),
    ('1.', '이자비용', 'Interest expenses', 2, None),
    ('2.', '외환차손', 'Losses Foreign Currency Transactions', 2, None),
    ('3.', '외화환산손실', 'Losses on foreign exchagne translations', 2, None),
    ('4.', '보증비용', '', 2, None),
    ('5.', '투자자산처분손실', 'Losses on disposal of investmens', 2, None),
    ('6.', '매도가능증권손상차손', 'Losses on valuation of available-for-sale financial assets', 2, None),
    ('7.', '기타금융자산평가손실', 'Losses on valuation of fair value financial asset', 2, None),
    ('8.', '기타금융자산처분손실', 'Losses on disposal of fair value financial asset', 2, None),
    ('9.', '매도가능증권처분손실', 'Losses on disposal of available-for-sale financial assets', 2, None),
    ('10.', '파생상품평가손실', 'Losses on valuation of derivative financial assets', 2, None),
    ('11.', '파생상품거래손실', 'Losses on disposal of derivative financial assets', 2, None),
    ('Ⅹ.', '지분법손익', '', 0, [(99, 1), (100, 1)]),
    ('1', '관계기업투자자산평가이익', '', 2, None),
    ('2', '관계기업투자자산평가손실', '', 2, None),
    ('XI.', '법인세비용차감전 계속사업이익', 'Profit (loss) before tax', 0, [(53, 1), (54, 1), (75, 1), (63, -1), (86, -1), (98, 1)]),
    ('XⅡ.', '법인세비용', 'Income tax expense', 0, None),
    ('ⅩⅢ.', '계속사업이익', 'Profit (loss) from continuing operations', 0, [(101, 1), (102, -1)]),
    ('XⅣ.', '중단사업이익', 'Profit (loss) from discontinued operations', 0, None),
    ('XⅤ.', '당기순이익(손실)', 'Profit (loss)', 0, [(103, 1), (104, -1)]),
    ('XⅥ.', '기타포괄손익', '', 0, None),
    ('가.', '당기손익재분류항목', '', 1, [(108, 1), (109, 1), (110, 1), (111, 1), (112, 1), (113, 1)]),
    ('1.', '기타포괄손익-기타금융자산평가이익', '', 2, None),
    ('2.', '기타포괄손익-기타금융자산평가손실', '', 2, None),
    ('3.', '지분법자본변동', '', 2, None),
    ('4.', '부의지분법자본변동', '', 2, None),
    ('5.', '해외사업환산이익', '', 2, None),
    ('6.', '해외사업환산손실', '', 2, None),
    ('나.', '당기손익비분류항목', '', 1, [(115, 1), (116, 1), (117, 1), (118, 1), (119, 1), (120, 1), (121, 1), (122, 1)]),
    ('1.', '재측정요소-이익', '', 2, None),
    ('2.', '재측정요소-손실', '', 2, None),
    ('3.', '기타포괄손익-기타금융자산평가이익', '', 2, None),
    ('4.', '기타포괄손익-기타금융자산평가손실', '', 2, None),
    ('5.', '지분법자본변동', '', 2, None),
    ('6.', '부의지분법자본변동', '', 2, None),
    ('7.', '해외사업환산이익', '', 2, None),
    ('8.', '해외사업환산손실', '', 2, None),
    ('XⅦ.', '총포괄손익', '', 0, [(105, 1), (107, 1), (114, 1)]),
    ('XⅧ.', '주당손익', '', 0, None),
    ('1.', '기본주당순손익', '', 2, None),
]

접이메뉴 = {'실적보고': 실적하위, '연결재무제표 패키지': 패키지하위}
보고서페이지 = ['1. Executive Summary', '2. P&L | 손익계산서', '3. SG&A | 판관비',
                '4. 운영 KPI', '5. Sales Forecast']
if 'rep_page' not in st.session_state:
    st.session_state['rep_page'] = 보고서페이지[0]

def _메뉴이동(n):
    # 상위 메뉴를 누르면 그 아래 첫 화면으로 갑니다 (실적보고 → 누적 실적보고)
    st.session_state['menu'] = 접이메뉴[n][0] if n in 접이메뉴 else n


def _보고서이동(n):
    st.session_state['menu'] = '월간실적'
    st.session_state['rep_page'] = n


with st.sidebar.container(key='nav_container'):
    st.caption('보기')
    for name in 메뉴목록:
        하위 = 접이메뉴.get(name, [])
        안쪽 = bool(하위) and st.session_state['menu'] in 하위
        펼침 = 안쪽 and len(하위) > 1        # 하위 화면이 하나뿐이면(실적보고) 펼치지 않습니다
        st.button(name, key=f'nav_{name}', width='stretch',
                  type='primary' if (st.session_state['menu'] == name or 안쪽) else 'secondary',
                  on_click=_메뉴이동, args=(name,))
        # 들여쓰기는 글자(공백) 대신 CSS 로 줍니다 — 브라우저 번역기가
        # 앞 공백을 지워도 줄이 어긋나지 않습니다
        if 펼침:
            with st.container(key='pl_sub'):
                for sp in 하위:
                    st.button(sp, key=f'pl_{sp}', width='stretch',
                              type='primary' if st.session_state['menu'] == sp else 'secondary',
                              on_click=_메뉴이동, args=(sp,))
        if name == '월간실적' and st.session_state['menu'] == '월간실적':
            with st.container(key='rep_sub'):
                for sp in 보고서페이지:
                    st.button(sp, key=f'rep_{sp}', width='stretch',
                              type='primary' if st.session_state['rep_page'] == sp else 'secondary',
                              on_click=_보고서이동, args=(sp,))

메뉴 = st.session_state['menu']
st.sidebar.write('---')

CACHE_PATH = '최근업로드.xlsx'
이름기록 = '최근업로드_파일이름.txt'
자료있음 = os.path.exists(CACHE_PATH)

# ══ 자료 보관소 — 웹에서도 자료가 사라지지 않게 ══════════════════════════
#    웹(Streamlit Cloud)은 서버가 잠들었다 깨어나면 올려 둔 파일을 지웁니다.
#    그래서 올려 주신 실적 엑셀을 「_보관자료」 폴더에 한 벌 넣고, 이 폴더를
#    GitHub 비공개 저장소에 함께 올려 둡니다. 대시보드는 열릴 때 그 파일을
#    그대로 쓰므로 아무도 올리지 않아도 늘 최신 자료가 보입니다.
#    ※ 저장소는 비공개이고, 넣어 두는 것은 「가장 최근 실적 엑셀 한 개」뿐입니다.
#    ※ 웹에서 올린 자료까지 저장소에 넣으려면 Streamlit 설정(Secrets)에
#      github_token 과 github_repo 두 줄이 있어야 합니다. 없으면 그 서버가
#      살아 있는 동안만 유지됩니다(예전과 같음).
보관폴더 = '_보관자료'
보관파일 = 보관폴더 + '/최신_실적엑셀.xlsx'
보관이름표 = 보관폴더 + '/최신_실적엑셀_이름.txt'


def _보관설정(열쇠):
    """Streamlit 설정(Secrets)에서 값을 꺼냅니다 — 없으면 빈 글자."""
    try:
        return str(st.secrets.get(열쇠, '') or '').strip()
    except Exception:
        return ''


보관KR파일 = 보관폴더 + '/최신_씨티케이커넥트.xlsx'
보관KR이름표 = 보관폴더 + '/최신_씨티케이커넥트_이름.txt'


def 보관소에서받기():
    """웹 서버가 새로 켜져 자료가 비었으면 GitHub 자료 저장소에서 두 법인 마지막 파일을 내려받습니다."""
    import urllib.request as _UR
    토큰, 저장소 = _보관설정('github_token'), _보관설정('github_repo')
    if not 토큰 or not 저장소:
        return
    def _받기(길):
        요청 = _UR.Request('https://api.github.com/repos/%s/contents/%s' % (저장소, 길),
                           headers={'Authorization': 'Bearer ' + 토큰, 'Accept': 'application/vnd.github.raw',
                                    'X-GitHub-Api-Version': '2022-11-28', 'User-Agent': 'ctk-dashboard'})
        with _UR.urlopen(요청, timeout=120) as 답:
            return 답.read()
    if not os.path.exists(보관용량자리):
        보관용량확인()
    for 파일, 이름표, 둘곳, 둘이름 in ((보관파일, 보관이름표, CACHE_PATH, 이름기록),
                                       (보관KR파일, 보관KR이름표, '최근업로드_KR.xlsx', '최근업로드_KR_파일이름.txt')):
        if os.path.exists(둘곳):
            continue
        try:
            값 = _받기(파일)
            with open(둘곳, 'wb') as f:
                f.write(값)
            try:
                with open(둘이름, 'wb') as f:
                    f.write(_받기(이름표))
            except Exception:
                pass
        except Exception:
            pass


def 보관소에저장(바이트, 파일이름, 법인자리='US'):
    """실적 엑셀을 GitHub 비공개 저장소에 넣습니다. 설정이 없으면 아무 일도 안 합니다."""
    import base64
    import json as _json
    import urllib.request as _UR

    토큰, 저장소 = _보관설정('github_token'), _보관설정('github_repo')
    if not 토큰 or not 저장소:
        return False
    머리 = {'Authorization': 'Bearer ' + 토큰,
            'Accept': 'application/vnd.github+json',
            'X-GitHub-Api-Version': '2022-11-28',
            'User-Agent': 'ctk-dashboard'}

    def _부르기(길, 몸=None, 방법='GET'):
        H = dict(머리)
        자료 = None
        if 몸 is not None:
            자료 = _json.dumps(몸).encode('utf-8')
            H['Content-Type'] = 'application/json'
        요청 = _UR.Request('https://api.github.com/repos/%s/%s' % (저장소, 길),
                           data=자료, headers=H, method=방법)
        with _UR.urlopen(요청, timeout=180) as 답:
            return _json.loads(답.read().decode('utf-8') or 'null')

    옛것 = {}
    try:
        목록 = _부르기('contents/' + 보관폴더)
        for e in (목록 if isinstance(목록, list) else []):
            옛것[e.get('path', '')] = e.get('sha', '')
    except Exception:
        pass

    def _올리기(길, 값, 메모):
        몸 = {'message': 메모, 'content': base64.b64encode(값).decode('ascii')}
        if 옛것.get(길):
            몸['sha'] = 옛것[길]
        _부르기('contents/' + 길, 몸, 'PUT')

    # 이번에 바꿀 파일들
    새것 = {}
    if 법인자리 == 'KR':      # 씨티케이커넥트 연결패키지
        새것[보관KR파일] = 바이트
        새것[보관KR이름표] = 파일이름.encode('utf-8')
    else:
        새것[보관파일] = 바이트
        새것[보관이름표] = 파일이름.encode('utf-8')
        try:                                   # 누가 · 언제 올렸는지도 함께
            with open(저장소기록자리, 'rb') as _f:
                새것[저장소기록자리] = _f.read()
        except Exception:
            pass

    # ① 과거 기록 없이 「지금 파일만 든 커밋」 하나로 바꿔 끼웁니다 (사용자 요청 : 과거자료는 지워도 됨 · 2026-11-03)
    try:
        가지 = _부르기('git/ref/heads/main')
        지금커밋 = _부르기('git/commits/' + 가지['object']['sha'])
        새트리 = [{'path': 길, 'mode': '100644', 'type': 'blob',
                   'sha': _부르기('git/blobs', {'content': base64.b64encode(값).decode('ascii'),
                                                'encoding': 'base64'}, 'POST')['sha']}
                  for 길, 값 in 새것.items()]
        트리 = _부르기('git/trees', {'base_tree': 지금커밋['tree']['sha'], 'tree': 새트리}, 'POST')
        커밋 = _부르기('git/commits', {'message': '최신 자료 — ' + 파일이름, 'tree': 트리['sha'], 'parents': []}, 'POST')
        _부르기('git/refs/heads/main', {'sha': 커밋['sha'], 'force': True}, 'PATCH')
    except Exception:
        # ② 저장소가 아직 비어 있는 등 ① 이 안 되면 예전 방식(파일마다 덮어쓰기)으로
        for 길, 값 in 새것.items():
            _올리기(길, 값, '자료 갱신 — ' + 파일이름)
    보관용량확인()
    return True


보관용량자리 = '최근보관용량.txt'
보관용량경고MB = 800          # GitHub 권장 1GB — 그 전에 알림


def 보관용량확인():
    """자료 저장소 용량(MB)을 GitHub 에 물어 적어 둡니다 — 사이드바 알림용."""
    import json as _json
    import urllib.request as _UR
    토큰, 저장소 = _보관설정('github_token'), _보관설정('github_repo')
    if not 토큰 or not 저장소:
        return None
    try:
        요청 = _UR.Request('https://api.github.com/repos/%s' % 저장소,
                           headers={'Authorization': 'Bearer ' + 토큰, 'Accept': 'application/vnd.github+json',
                                    'X-GitHub-Api-Version': '2022-11-28', 'User-Agent': 'ctk-dashboard'})
        with _UR.urlopen(요청, timeout=30) as 답:
            MB = float(_json.loads(답.read().decode('utf-8')).get('size', 0)) / 1024
        with open(보관용량자리, 'w', encoding='utf-8') as f:
            f.write(f'{MB:.1f}')
        return MB
    except Exception:
        return None


def 보관용량():
    try:
        return float(open(보관용량자리, encoding='utf-8').read().strip())
    except Exception:
        return None


def 보관해두기(바이트, 파일이름):
    """올려 주신 실적 엑셀을 보관소에 넣습니다 (같은 파일이면 한 번만)."""
    도장 = (str(파일이름), len(바이트))
    if st.session_state.get('_보관도장') == 도장:
        return
    st.session_state['_보관도장'] = 도장
    try:
        os.makedirs(보관폴더, exist_ok=True)
        with open(보관파일, 'wb') as _f:
            _f.write(바이트)
        with open(보관이름표, 'w', encoding='utf-8') as _f:
            _f.write(str(파일이름))
    except OSError:
        pass
    try:                                   # 누가 · 언제 올렸는지 남깁니다
        자료기록쓰기(파일이름, 올린이이름())
    except Exception:
        pass
    try:
        with st.spinner('자료를 보관소에 저장하고 있습니다 — 잠시만요'):
            st.session_state['_보관결과'] = '됨' if 보관소에저장(바이트, 파일이름) else ''
    except Exception as _e:
        st.session_state['_보관결과'] = '안됨:' + str(_e)[:120]


기록자리 = os.path.join(os.path.dirname(CACHE_PATH) or '.', '최근업로드_기록.txt')
저장소기록자리 = 보관폴더 + '/최신_실적엑셀_기록.txt'      # 저장소에 함께 올라갑니다


def 지금시각():
    """보고 계신 분 시계 기준 「지금」 (웹서버는 영국 시간이라 그대로 쓰면 어긋납니다)."""
    try:
        이름 = getattr(st.context, 'timezone', None)
        if 이름:
            import zoneinfo
            return datetime.datetime.now(zoneinfo.ZoneInfo(str(이름)))
    except Exception:
        pass
    try:
        분 = getattr(st.context, 'timezone_offset', None)
        if 분 is not None:
            return (datetime.datetime.now(datetime.timezone.utc)
                    - datetime.timedelta(minutes=int(분)))
    except Exception:
        pass
    return datetime.datetime.now()


def 올린이이름():
    """올린 곳 — 「실적보고 엑셀작성」 화면에서 고른 것이 있으면 그것, 없으면 사이드바 설정칸."""
    return str(st.session_state.get('올린곳_작성') or st.session_state.get('올린이', '') or '').strip()


def 자료기록읽기():
    """{'이름': 파일명, '올린이': 사람, '때': '2026-09-23 14:20'} — 없으면 {}."""
    for 자리 in (기록자리, 저장소기록자리):
        try:
            with open(자리, encoding='utf-8') as f:
                d = json.load(f)
            if isinstance(d, dict):
                return d
        except Exception:
            continue
    return {}


def 자료기록쓰기(이름, 누가):
    쪽지 = {'이름': 이름 or '', '올린이': 누가 or '',
            '때': f'{지금시각():%Y-%m-%d %H:%M}'}
    for 자리 in (기록자리, 저장소기록자리):
        try:
            os.makedirs(os.path.dirname(자리) or '.', exist_ok=True)
            with open(자리, 'w', encoding='utf-8') as f:
                json.dump(쪽지, f, ensure_ascii=False)
        except Exception:
            pass
    return 쪽지


# 웹 서버가 새로 켜졌으면 GitHub 자료 저장소에서 두 법인 마지막 파일을 먼저 내려받습니다 (2026-11-03)
if not 자료있음 or not os.path.exists('최근업로드_KR.xlsx'):
    try:
        보관소에서받기()
    except Exception:
        pass
    자료있음 = os.path.exists(CACHE_PATH)
# 서버가 다시 켜져 캐시가 비었어도, 보관해 둔 자료로 바로 화면을 그립니다
if not 자료있음 and os.path.exists(보관파일):
    try:
        import shutil as _shutil
        _둘곳 = os.path.dirname(CACHE_PATH)
        if _둘곳:
            os.makedirs(_둘곳, exist_ok=True)
        _shutil.copyfile(보관파일, CACHE_PATH)
        if os.path.exists(보관이름표):
            _shutil.copyfile(보관이름표, 이름기록)
        자료있음 = True
    except OSError:
        pass
# 웹사이트를 업데이트(3번 bat)해 보관 자료가 더 새것이면 그것으로 바꿉니다 — 예전에 웹에 올린 파일이 남아 새 양식이 안 보이는 일 방지 (2026-10-27)
elif os.path.exists(보관파일) and os.path.exists(CACHE_PATH):
    try:
        if os.path.getmtime(보관파일) > os.path.getmtime(CACHE_PATH) + 60:
            with open(보관파일, 'rb') as _a, open(CACHE_PATH, 'rb') as _b:
                _다름 = _a.read() != _b.read()
            if _다름:
                import shutil as _shutil
                _shutil.copyfile(보관파일, CACHE_PATH)
                if os.path.exists(보관이름표):
                    _shutil.copyfile(보관이름표, 이름기록)
    except OSError:
        pass

# CTKonnect (2026-10-29) — 자료실 달 폴더(예: 자료실\\2026년 8월)에 넣어 둔 연결패키지 중 가장 새 것이
#   지금 쓰는 자료보다 새로우면 그것으로 바꿉니다. 매달 패키지 파일만 폴더에 넣으면 됩니다.
def _자료실최신패키지():
    import glob as _glob
    후보 = []
    for f in _glob.glob(os.path.join('자료실', '**', '*.xlsx'), recursive=True):
        n = os.path.basename(f)
        if n.startswith('~$') or '연결패키지' not in n or '씨티케이커넥트' in n:
            continue
        try:
            with zipfile.ZipFile(f) as z:
                w = z.read('xl/workbook.xml').decode('utf-8', 'ignore')
        except Exception:
            continue
        if '원장_raw' in w:
            후보.append(f)
    return max(후보, key=os.path.getmtime) if 후보 else None


try:
    _최신 = _자료실최신패키지()
    if _최신 and (not os.path.exists(CACHE_PATH)
                  or os.path.getmtime(_최신) > os.path.getmtime(CACHE_PATH) + 1):
        import shutil as _shutil
        _shutil.copyfile(_최신, CACHE_PATH)
        os.utime(CACHE_PATH, (os.path.getmtime(_최신), os.path.getmtime(_최신)))
        with open(이름기록, 'w', encoding='utf-8') as _f:
            _f.write(os.path.basename(_최신))
        자료있음 = True
except OSError:
    pass

# 한 번 올리면 그대로 남겨 두는 자료 — 다음에 또 올리지 않아도 됩니다
원장보관 = '최근원장.xlsx'            # ① 퀵북 원장 원본
통장보관 = '최근통장.xlsx'            # ② 2026년 금융거래 (CTKonnect)
전월보관 = '최근전월실적.xlsx'        # ② 지난달 실적보고 엑셀


def 자료남기기(자리, 올림):
    """올려 주신 파일을 그대로 남겨 둡니다 (같은 파일이면 다시 쓰지 않습니다)."""
    if 올림 is None:
        return
    도장 = f'{올림.name}|{올림.size}'
    if st.session_state.get('보관도장_' + 자리) == 도장 and os.path.exists(자리):
        return
    try:
        with open(자리, 'wb') as f:
            f.write(올림.getvalue())
        with open(자리 + '.txt', 'w', encoding='utf-8') as f:
            f.write(올림.name)
        st.session_state['보관도장_' + 자리] = 도장
    except OSError:
        pass


def 보관자료(자리):
    """(있나, 원래 파일이름, 크기, 남긴 때) — 없으면 (False, '', 0, 0)."""
    try:
        s = os.stat(자리)
    except OSError:
        return False, '', 0, 0
    try:
        with open(자리 + '.txt', encoding='utf-8') as f:
            이름 = f.read().strip()
    except OSError:
        이름 = 자리
    return True, 이름, s.st_size, s.st_mtime


def 보관바이트(자리):
    try:
        with open(자리, 'rb') as f:
            return f.read()
    except OSError:
        return None


def 보관원장정보():
    """(있나, 파일이름, 남긴 때) — 예전 이름 그대로 쓰는 곳을 위해 남겨 둡니다."""
    있음, 이름, _크기, 때 = 보관자료(원장보관)
    return 있음, 이름, 때


def _자리에쓰기(자리, 바이트, 이름):
    try:
        with open(자리, 'wb') as f:
            f.write(바이트)
        with open(자리 + '.txt', 'w', encoding='utf-8') as f:
            f.write(이름)
        return True
    except OSError:
        return False


from openpyxl.styles import Font
# ══════════════════════════════════════════════════════════════════════════
# 씨티케이커넥트(한국 법인) 연결패키지 만들기 — 2026-11-03
#   더존 분개장(「분개장_…」 시트) + 전월 씨티케이커넥트 연결패키지 → 결산월 연결패키지
#   · 전월 패키지를 그대로 두고 「26년 원장_raw」 만 새 분개장으로 갈아 끼웁니다 (BS · IS · 요약보고 수식 · 양식은 그대로)
#   · 기초잔액 = 전월 패키지 「26년 BS기초」(25년 말 확정 재무상태표)
#   · 계정 연결 = 전월 패키지 「BS_IS_매핑」 — 새 계정은 계정번호로 임시 연결하고 빨간 「확인 필요」
#   · Q 활동분류 · R 활동세부 = 전월 원장_raw 의 같은 전표 줄에서 이어받음 (없으면 계정 기준 기본값)
# ══════════════════════════════════════════════════════════════════════════
def 한국분개장인지(바이트):
    """더존 분개장 파일인지 — 「분개장」 시트가 있고 머리글에 계정과목 · 차변 · 대변."""
    try:
        _w = openpyxl.load_workbook(io.BytesIO(바이트), read_only=True)
        for n in _w.sheetnames:
            if '분개장' in n:
                for r in _w[n].iter_rows(min_row=1, max_row=8, values_only=True):
                    글 = ' '.join(str(x) for x in r if x is not None)
                    if '계정과목' in 글 and '차변' in 글 and '대변' in 글:
                        return True
    except Exception:
        pass
    return False


_KR기본과목 = {
    '103': ('BS', '현금및현금성자산'), '108': ('BS', '매출채권'), '120': ('BS', '미수금'), '122': ('BS', '기타금융자산(유동)'),
    '131': ('BS', '선급금'), '133': ('BS', '선급비용'), '134': ('BS', '미수금'), '135': ('BS', '부가세대급금'),
    '136': ('BS', '당기법인세자산'), '146': ('BS', '상품'), '150': ('BS', '제품'), '169': ('BS', '재공품'),
    '171': ('BS', '재공품'), '178': ('BS', '기타금융자산(비유동)'), '182': ('BS', '관계기업투자자산'),
    '198': ('BS', '사외적립자산'), '212': ('BS', '비품'), '213': ('BS', '비품-감가상각누계액'), '214': ('BS', '건설중인자산'),
    '233': ('BS', '상표권'), '240': ('BS', '소프트웨어'), '251': ('BS', '매입채무'), '253': ('BS', '미지급금'),
    '254': ('BS', '예수금'), '255': ('BS', '부가세예수금'), '259': ('BS', '선수금'), '260': ('BS', '단기차입금'),
    '262': ('BS', '미지급비용'), '276': ('BS', '전환사채(유동)'), '277': ('BS', '전환사채(유동)'), '278': ('BS', '전환사채(유동)'),
    '279': ('BS', '파생상품금융부채(유동)'), '295': ('BS', '확정급여채무'), '331': ('BS', '보통주자본금'),
    '341': ('BS', '주식발행초과금'), '349': ('BS', '기타자본잉여금'), '377': ('BS', '미처분이익잉여금'),
    '381': ('BS', '주식할인발행차금'), '384': ('BS', '기타자본조정'), '397': ('BS', '주식매수선택권'),
    '401': ('IS', '상품매출'), '404': ('IS', '제품매출'), '414': ('IS', '용역매출'), '451': ('IS', '상품매출원가'),
    '455': ('IS', '제품매출원가'), '457': ('IS', '용역매출원가'), '533': ('IS', '제품매출원가'), '537': ('IS', '제품매출원가'),
    '733': ('IS', '용역매출원가'), '801': ('IS', '급여'), '802': ('IS', '급여'), '808': ('IS', '퇴직급여'),
    '811': ('IS', '복리후생비'), '812': ('IS', '여비교통비'), '813': ('IS', '접대비'), '814': ('IS', '통신비'),
    '815': ('IS', '수도광열비'), '816': ('IS', '전력비'),
    '817': ('IS', '세금과공과'), '818': ('IS', '감가상각비'), '819': ('IS', '지급임차료'), '820': ('IS', '수선비'),
    '821': ('IS', '보험료'), '822': ('IS', '차량유지비'),
    '823': ('IS', '경상연구개발비'), '824': ('IS', '운반비'), '825': ('IS', '교육훈련비'), '826': ('IS', '도서인쇄비'),
    '827': ('IS', '잡비'), '829': ('IS', '사무용품비'), '830': ('IS', '소모품비'), '831': ('IS', '지급수수료'),
    '833': ('IS', '광고선전비'), '835': ('IS', '대손상각비'), '840': ('IS', '무형자산상각비'), '842': ('IS', '견본비'),
    '901': ('IS', '이자수익'), '903': ('IS', '배당금수익'), '907': ('IS', '외환차손익'), '910': ('IS', '외화환산손익'),
    '930': ('IS', '잡이익'), '931': ('IS', '이자비용'), '932': ('IS', '외환차손익'), '933': ('IS', '기부금'),
    '935': ('IS', '외화환산손익'), '960': ('IS', '잡손실'), '998': ('IS', '법인세비용')}
_KR기본활동 = {'801': '인건비성 항목', '802': '인건비성 항목', '808': '인건비성 항목', '811': '인건비성 항목',
              '819': '건물', '814': 'ICT', '812': '출장비', '824': '운반 및 창고료', '831': '외부용역',
              '813': '소모성경비', '826': '소모성경비', '830': '소모성경비', '842': '소모성경비',
              '827': '소모성경비', '825': '소모성경비', '829': '소모성경비', '822': '차량'}
_KR수익과목 = {'상품매출', '제품매출', '용역매출', '기타매출', '패키지매출', '이자수익', '잡이익', '외환차손익',
              '외화환산손익', '배당금수익'}


def _KR분개장읽기(바이트):
    xls = pd.ExcelFile(io.BytesIO(바이트))
    시트 = next(n for n in xls.sheet_names if '분개장' in n)
    raw = pd.read_excel(xls, sheet_name=시트, header=None, nrows=12)
    머리 = next(i for i in range(len(raw))
               if any(str(x).replace('\n', '') == '계정과목' for x in raw.iloc[i].tolist()))
    d = pd.read_excel(xls, sheet_name=시트, header=머리)
    d.columns = [str(c).replace('\n', '').strip() for c in d.columns]
    d = d[d['일자'].map(lambda x: hasattr(x, 'year') or bool(re.match(r'^\d{4}-\d{2}-\d{2}', str(x))))].copy()
    d['일자'] = pd.to_datetime(d['일자'])
    d['코드'] = pd.to_numeric(d['계정코드'], errors='coerce').fillna(0).astype(int).astype(str)
    d['계정'] = d['코드'] + ' ' + d['계정과목'].astype(str).str.strip()
    d['차'] = pd.to_numeric(d['차변'], errors='coerce').fillna(0.0)
    d['대'] = pd.to_numeric(d['대변'], errors='coerce').fillna(0.0)
    d['전표'] = (d['일자'].dt.strftime('%m%d') + '-'
                 + pd.to_numeric(d['전표번호'], errors='coerce').fillna(0).astype(int).astype(str).str.zfill(5))
    for c in ('적요', '거래처명', '유형'):
        if c not in d.columns:
            d[c] = None
    if abs(d['차'].sum() - d['대'].sum()) >= 1:
        raise ValueError(f'분개장 차변 합계와 대변 합계가 다릅니다 ({d["차"].sum():,.0f} / {d["대"].sum():,.0f})')
    return d, 시트


def KR패키지만들기(분개장바이트, 전월바이트):
    """→ (엑셀 바이트, 안내 dict). 전월바이트 = 전월(또는 같은 달 이전 버전) 씨티케이커넥트 연결패키지."""
    RED = Font(name='맑은 고딕', size=9, bold=True, color='FFC00000')
    d, 시트이름 = _KR분개장읽기(분개장바이트)
    끝날 = d['일자'].max()
    연도, 결산월 = int(끝날.year), int(끝날.month)
    wb = openpyxl.load_workbook(io.BytesIO(전월바이트))
    if '26년 원장_raw' not in wb.sheetnames or '씨티케이커넥트' not in str(wb['26년 원장_raw']['A1'].value or ''):
        raise ValueError('③ 전월 연결패키지가 씨티케이커넥트 패키지가 아닙니다 — 「…연결패키지_씨티케이커넥트_…」 파일을 올려 주세요.')
    안내 = {'연도': 연도, '결산월': 결산월, '줄': len(d), '새계정': [], '알림': []}

    # ── BS_IS_매핑 : 있는 것 그대로 + 새 계정 추가
    M = wb['BS_IS_매핑']
    매핑 = {}
    마지막 = 1
    for r in range(2, M.max_row + 1):
        a = M.cell(r, 1).value
        if a:
            매핑[str(a).strip()] = (M.cell(r, 2).value, M.cell(r, 3).value)
            마지막 = r
    틀 = [_커_스타일(M.cell(2, c)) for c in range(1, 5)]
    for 계정 in sorted(set(d['계정']) - set(매핑), key=lambda s: int(re.match(r'^\d+', s).group(0))):
        c = 계정.split(' ', 1)[0]
        k, g = _KR기본과목.get(c, ('BS' if c[0] in '123' else 'IS',
                                   '잡비' if c[0] == '8' else ('잡손실' if c[0] == '9' else '기타')))
        마지막 += 1
        for j, v in enumerate((계정, k, g, '새 계정 — 계정번호로 임시 연결 (확인 필요)'), 1):
            _커_입히기(M.cell(마지막, j), 틀[min(j - 1, 3)])
            M.cell(마지막, j).value = v
        M.cell(마지막, 4).font = RED
        매핑[계정] = (k, g)
        안내['새계정'].append(f'{계정} → {g}')

    # ── 기초잔액 : 26년 BS기초 (A 계정명 · B 25년 말)
    B = wb['26년 BS기초']
    기초 = {}
    for r in range(7, B.max_row + 1):
        a, v = B.cell(r, 1).value, B.cell(r, 2).value
        if a and isinstance(v, (int, float)):
            기초[str(a).strip()] = float(v)

    # ── 자연 잔액 방향 (자산 · 비용 = 차변)
    _자산과목 = {str(wb['BS'].cell(rr, 5).value or '').replace(' ', '') for rr in range(9, 93)}

    def 차변성(계정):
        k, g = 매핑.get(계정, ('BS', ''))
        g = str(g or '').replace(' ', '')
        return (g in _자산과목) if k == 'BS' else (g not in _KR수익과목)

    # ── 전월 원장_raw 에서 손으로 넣은 Q · R · S · T 이어받기 (전표 · 계정 · 차 · 대 가 같은 줄)
    W = wb['26년 원장_raw']
    이어 = {}
    for row in W.iter_rows(min_row=6, values_only=True):
        if row[2] is not None and row[1]:
            키 = (str(row[5]), str(row[1]).strip(), round(float(row[8] or 0), 0), round(float(row[9] or 0), 0))
            이어[키] = row[16:20] if len(row) >= 20 else (None,) * 4

    # ── 원장_raw 다시 쓰기 (머리 · 기초 · 거래 · 합계 줄의 서식 · 수식은 전월 줄에서)
    tot = next(r for r in range(7, W.max_row + 1) if str(W.cell(r, 1).value or '').startswith('Total for'))
    본 = {'머리': 6, '기초': 7, '거래': 8, '합계': tot}
    모양 = {k: {c: _커_스타일(W.cell(r, c)) for c in range(1, 23)} for k, r in 본.items()}
    수식 = {k: {c: W.cell(r, c).value for c in range(11, 23)
               if isinstance(W.cell(r, c).value, str) and W.cell(r, c).value.startswith('=')} for k, r in 본.items()}
    for rng in list(W.merged_cells.ranges):
        if rng.min_row >= 6:
            W.unmerge_cells(str(rng))
    W.delete_rows(6, W.max_row)
    W['A3'] = f'{연도}.01.01 ~ {끝날:%Y.%m.%d}   (F = 전표번호 월일-번호 · G = 전표 유형)'

    def 쓰기(r, 종류, 값):
        for c in range(1, 23):
            _커_입히기(W.cell(r, c), 모양[종류].get(c))
        for c, f in 수식[종류].items():
            W.cell(r, c, _커_옮김(f, 본[종류], r, openpyxl.utils.get_column_letter(c)))
        for c, v in 값.items():
            W.cell(r, c, v)

    이어받음 = 0
    r = 6
    for 계정 in sorted(set(d['계정']), key=lambda s: int(re.match(r'^\d+', s).group(0))):
        c = 계정.split(' ', 1)[0]
        x = d[d['계정'] == 계정].sort_values(['일자', '전표'], kind='stable')
        쓰기(r, '머리', {1: 계정}); r += 1
        잔 = float(기초.get(계정, 0.0)); 차성 = 차변성(계정)
        쓰기(r, '기초', {2: 'Beginning Balance', 8: 잔}); r += 1
        첫 = r
        for _, y in x.iterrows():
            잔 += (y['차'] - y['대']) if 차성 else (y['대'] - y['차'])
            값 = {2: 계정, 3: y['일자'].to_pydatetime(),
                  4: str(y['적요']) if pd.notna(y['적요']) else None,
                  5: str(y['거래처명']) if pd.notna(y['거래처명']) else None,
                  6: y['전표'], 7: str(y['유형']) if pd.notna(y['유형']) else None,
                  8: round(잔, 0), 9: (y['차'] or None), 10: (y['대'] or None)}
            옛 = 이어.get((y['전표'], 계정, round(y['차'], 0), round(y['대'], 0)))
            if 옛 and any(v is not None for v in 옛):
                for j, v in enumerate(옛):
                    if v is not None:
                        값[17 + j] = v
                이어받음 += 1
            elif c in _KR기본활동:
                값[17] = _KR기본활동[c]; 값[18] = 계정.split(' ', 1)[1].replace('(판)', '')
            쓰기(r, '거래', 값); r += 1
        쓰기(r, '합계', {1: f'Total for {계정}', 8: round(잔, 0),
                         9: f'=SUM(I{첫}:I{r - 1})', 10: f'=SUM(J{첫}:J{r - 1})'}); r += 1
    W.auto_filter.ref = f'A5:V{r - 1}'
    안내['이어받음'] = 이어받음

    # ── 머리 글자 (기간)
    말일 = f'{연도}년 {결산월:02d}월 {끝날.day:02d}일'
    bs, is_ = wb['BS'], wb['IS']
    bs['B2'] = re.sub(r'\d{4}년\s*\d{1,2}월\s*\d{1,2}일', 말일, str(bs['B2'].value or '')) or f'(당) : {말일} 현재'
    is_['B2'] = re.sub(r'~\s*\d{4}\.\d{2}\.\d{2}', f'~{끝날:%Y.%m.%d}', str(is_['B2'].value or ''))
    if '요약보고' in wb.sheetnames:
        S = wb['요약보고']
        for row in S.iter_rows(min_row=1, max_row=8):
            for x in row:
                if isinstance(x.value, str) and re.fullmatch(r'\d{1,2}월 누적', x.value.strip()):
                    x.value = f'{결산월}월 누적'
        # 3 · 4번 표에 없는 새 거래처 · 계정 — 표 합계 줄 비고가 2번 표와 차이를 보여 줌
        있는 = {str(S.cell(rr, 2).value or '').strip() for rr in range(1, S.max_row + 1)}
        _m = d[d['코드'].isin(['401', '404', '414'])]
        새거래처 = sorted({str(n).strip() for n in _m['거래처명'].dropna()} - 있는)
        _p = d[d['코드'].str.startswith('8')]
        새과목 = sorted({str(매핑.get(n, ('', ''))[1]) for n in _p['계정']} - 있는)
        if 새거래처:
            안내['알림'].append('요약보고 3번 표에 없는 매출 거래처 : ' + ', '.join(새거래처) + ' — 줄을 추가해 주세요 (합계 줄 비고에 차이로 표시됨)')
        if 새과목:
            안내['알림'].append('요약보고 4번 표에 없는 판관비 과목 : ' + ', '.join(새과목) + ' — 줄을 추가해 주세요')
    for w in wb:
        w.sheet_view.showGridLines = False
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), 안내


def 자료종류(바이트):
    """파일 안 시트 이름을 보고 어떤 자료인지 스스로 알아냅니다."""
    try:
        시트 = [str(s) for s in pd.ExcelFile(io.BytesIO(바이트)).sheet_names]
    except Exception:
        return ''
    if any('원장_raw' in s for s in 시트):
        return '실적'                       # 실적보고 엑셀 (대시보드가 그리는 자료)
    if any(s.strip() in ('BS-당기', 'IS-당기') for s in 시트):
        return '양식'                       # 본사 연결패키지 원본
    if any(re.search(r'\d+\s*월\s*\(', s) for s in 시트):
        return '원장'                       # 퀵북 원장 원본
    if any(re.search(r'ledger|원장', s, re.I) for s in 시트):
        return '원장'
    if any('분개장' in s for s in 시트) and 한국분개장인지(바이트):
        return '원장'                       # 씨티케이커넥트 더존 분개장 (2026-11-03)
    try:                                    # 퀵북 General Ledger — 머리글로 판별 (시트 이름이 「분개장_8월누적」 등이어도)
        _w = openpyxl.load_workbook(io.BytesIO(바이트), read_only=True)
        for _r in _w[_w.sheetnames[0]].iter_rows(min_row=1, max_row=8, values_only=True):
            _글 = ' '.join(str(x) for x in _r if x is not None)
            if 'Distribution account' in _글 and 'Transaction date' in _글:
                return '원장'
    except Exception:
        pass
    return ''


종류이름 = {'원장반영': 'CTKonnect 원장 → 패키지에 반영', '실적': 'CTKonnect 실적 엑셀', '한국실적': '씨티케이커넥트 연결패키지', '한국분개장': '씨티케이커넥트 분개장 → 패키지로 변환', '원장': '원장 원본', '양식': '연결패키지 양식'}


KR보관, KR이름표 = '최근업로드_KR.xlsx', '최근업로드_KR_파일이름.txt'      # 씨티케이커넥트(한국 법인) 연결패키지 자리


def 한국법인자료(바이트, 이름=''):
    """씨티케이커넥트(한국 법인) 연결패키지인지 — 파일 이름이나 원장_raw A1 의 회사 이름으로 봅니다."""
    if '씨티케이커넥트' in str(이름):
        return True
    try:
        _w = openpyxl.load_workbook(io.BytesIO(바이트), read_only=True)
        for _n in _w.sheetnames:
            if '원장_raw' in _n:
                _a1 = str(_w[_n]['A1'].value or '')
                return '씨티케이커넥트' in _a1
    except Exception:
        pass
    return False


def _KR패키지인가(바이트):
    """연결패키지(「원장_raw」 시트가 든 파일)인지."""
    try:
        with zipfile.ZipFile(io.BytesIO(바이트)) as z:
            return '원장_raw' in z.read('xl/workbook.xml').decode('utf-8', 'ignore')
    except Exception:
        return False


def KR분개장반영(바이트, 이름='', 양식=None, 양식이름=''):
    """더존 분개장 → 씨티케이커넥트 패키지(함께 올린 것, 없으면 지금 있는 것)를 양식으로 새 패키지를 만들어 KR 자리에 둡니다 → (됨, 안내글)."""
    도장 = (str(이름), len(바이트), str(양식이름), len(양식 or b''))
    if st.session_state.get('_KR분개장도장') == 도장:
        return True, st.session_state.get('_KR분개장안내', '')
    if 양식:
        전월, 전월이름 = 양식, 양식이름
    else:
        전월, 전월이름 = _KR패키지바이트(분개장보기=False)
    if not 전월:
        return False, ('분개장만으로는 처음 패키지를 만들 수 없습니다 — 씨티케이커넥트 연결패키지(「…연결패키지_씨티케이커넥트_…」)를 '
                       '한 번 올려 주시면, 그다음부터는 분개장만 올려도 됩니다.')
    with st.spinner('분개장으로 씨티케이커넥트 패키지를 만드는 중입니다 (20~40초)...'):
        결과, 안내 = KR패키지만들기(바이트, 전월)
    연, 월 = 안내['연도'], 안내['결산월']
    _m = re.search(rf'_{월}월말_v(\d+)', str(전월이름 or ''))
    새이름 = f'{연}. {(월 - 1) // 3 + 1}Q_연결패키지_씨티케이커넥트_{월}월말_v{int(_m.group(1)) + 1 if _m else 1}.xlsx'
    KR보관하기(결과, 새이름)
    글 = (f'분개장 {안내["줄"]:,}줄 → 「{새이름}」 (양식 = {전월이름})'
          + (f' · 새 계정 {len(안내["새계정"])}개는 BS_IS_매핑 확인 필요' if 안내['새계정'] else ''))
    st.session_state['_KR분개장도장'] = 도장
    st.session_state['_KR분개장안내'] = 글
    return True, 글


def KR보관하기(바이트, 이름):
    """씨티케이커넥트 패키지를 이 서버와 GitHub 자료 저장소에 남깁니다 (같은 파일이면 한 번만)."""
    if not _KR패키지인가(바이트):          # 분개장 등 패키지가 아닌 파일은 두지 않음
        return
    도장 = (str(이름), len(바이트))
    if st.session_state.get('_KR보관도장') == 도장:
        return
    st.session_state['_KR보관도장'] = 도장
    try:
        with open(KR보관, 'wb') as _f:
            _f.write(바이트)
        with open(KR이름표, 'w', encoding='utf-8') as _f:
            _f.write(str(이름))
    except OSError:
        pass
    try:
        with st.spinner('씨티케이커넥트 자료를 보관소에 저장하고 있습니다 — 잠시만요'):
            보관소에저장(바이트, str(이름), 'KR')
    except Exception:
        pass


def _KR패키지바이트(분개장보기=True):
    """씨티케이커넥트 연결패키지 → (바이트, 파일 이름). 방금 올린 것 → 자료실 가장 새 파일 → 마지막 보관 순."""
    _up = st.session_state.get('kr_pkg')
    if _up is not None:
        _b = _up.getvalue()
        if _KR패키지인가(_b):
            KR보관하기(_b, _up.name)
            return _b, _up.name
        if 분개장보기 and 한국분개장인지(_b):        # 분개장을 올리셨으면 패키지로 바꿔 KR 자리에 둠
            try:
                _됨, _글 = KR분개장반영(_b, _up.name)
            except Exception as e:
                _됨, _글 = False, f'분개장으로 패키지를 만들지 못했습니다 — {e}'
            st.session_state['_KR올림안내'] = (_됨, _글)
    후보 = [f for f in glob.glob(os.path.join(자료실, '**', '*.xlsx'), recursive=True)
            if '씨티케이커넥트' in os.path.basename(f) and '연결패키지' in os.path.basename(f)
            and not os.path.basename(f).startswith('~$')]
    _p = max(후보, key=os.path.getmtime) if 후보 else None
    if _p and (not os.path.exists(KR보관) or os.path.getmtime(_p) > os.path.getmtime(KR보관) + 1):
        try:
            import shutil as _sh
            _sh.copyfile(_p, KR보관)
            os.utime(KR보관, (os.path.getmtime(_p), os.path.getmtime(_p)))
            with open(KR이름표, 'w', encoding='utf-8') as _f:
                _f.write(os.path.basename(_p))
        except OSError:
            pass
        with open(_p, 'rb') as _f:
            return _f.read(), os.path.basename(_p)
    if os.path.exists(KR보관):
        with open(KR보관, 'rb') as _f:
            _b = _f.read()
        if not _KR패키지인가(_b):             # 예전에 분개장이 잘못 들어간 경우
            return None, ''
        try:
            _n = open(KR이름표, encoding='utf-8').read().strip()
        except OSError:
            _n = KR보관
        return _b, _n
    return None, ''


def _US원장인가(바이트):
    """퀵북 원장(General Ledger)인지 — 씨티케이커넥트 분개장 · 연결패키지는 아님."""
    if _KR패키지인가(바이트) or 한국분개장인지(바이트):
        return False
    return 자료종류(바이트) == '원장'


def US올린것처리(파일들):
    """CTKonnect 퀵북 원장 · 연결패키지를 한꺼번에 → 새 패키지를 만들어 화면 자료로 씁니다 → (됨, 안내글) 은 session 에."""
    패키지 = [f for f in 파일들 or [] if _KR패키지인가(f.getvalue()) and not 한국법인자료(f.getvalue(), f.name)]
    원장 = [f for f in 파일들 or [] if _US원장인가(f.getvalue())]
    도장 = tuple(sorted((f.name, len(f.getvalue())) for f in 패키지 + 원장))
    if not 도장 or st.session_state.get('_US올림도장') == 도장:
        return
    st.session_state['_US올림도장'] = 도장
    if 패키지:
        양식, 양식이름 = 패키지[-1].getvalue(), 패키지[-1].name
    elif os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, 'rb') as _f:
            양식 = _f.read()
        양식이름 = 쓰던파일이름() or CACHE_PATH
    else:
        양식, 양식이름 = None, ''
    if not 원장:
        결과, 새이름, 글 = 양식, 양식이름, f'CTKonnect 연결패키지 「{양식이름}」 를 그대로 씁니다.'
    elif not 양식:
        st.session_state['_US올림안내'] = (False, '원장만으로는 처음 패키지를 만들 수 없습니다 — CTKonnect 연결패키지를 함께 올려 주세요.')
        return
    else:
        _자리에쓰기(원장보관, 원장[-1].getvalue(), 원장[-1].name)
        try:
            with st.spinner('원장으로 CTKonnect 연결패키지를 만드는 중입니다 (10~20초)...'):
                결과, 안내 = 커_패키지만들기(원장[-1].getvalue(), 양식)
        except Exception as e:
            st.session_state['_US올림안내'] = (False, f'원장으로 패키지를 만들지 못했습니다 — {e}')
            return
        연, 월 = int(안내['연도']), int(안내['결산월'])
        _m = re.search(rf'_{월}월말_v(\d+)', str(양식이름))
        새이름 = f'{연}. {(월 - 1) // 3 + 1}Q_연결패키지_CTKonnect_{월}월말_v{int(_m.group(1)) + 1 if _m else 1}.xlsx'
        글 = (f'원장 「{원장[-1].name}」 → 「{새이름}」 (양식 = {양식이름}) · '
              f'{안내["원장끝"]:%Y-%m-%d} 까지 · 결산월 {월}월 · 환율은 양식 값 그대로')
    try:
        with open(CACHE_PATH, 'wb') as _f:
            _f.write(결과)
        with open(이름기록, 'w', encoding='utf-8') as _f:
            _f.write(새이름)
    except OSError:
        pass
    보관해두기(결과, 새이름)
    st.session_state['_US올림안내'] = (True, 글)


def KR올린것처리(파일들):
    """씨티케이커넥트 자료 여러 개(분개장 · 연결패키지)를 한꺼번에 → KR 자리에 반영하고 안내글을 남깁니다."""
    패키지 = [f for f in 파일들 or [] if _KR패키지인가(f.getvalue()) and 한국법인자료(f.getvalue(), f.name)]
    분개장 = [f for f in 파일들 or [] if 한국분개장인지(f.getvalue())]
    양식, 양식이름 = (패키지[-1].getvalue(), 패키지[-1].name) if 패키지 else (None, '')
    if 분개장:
        try:
            _됨, _글 = KR분개장반영(분개장[-1].getvalue(), 분개장[-1].name, 양식, 양식이름)
        except Exception as e:
            _됨, _글 = False, f'분개장으로 패키지를 만들지 못했습니다 — {e}'
        st.session_state['_KR올림안내'] = (_됨, _글)
    elif 양식:
        KR보관하기(양식, 양식이름)
        st.session_state['_KR올림안내'] = (True, f'씨티케이커넥트 연결패키지 「{양식이름}」 를 그대로 씁니다.')
    elif 파일들:
        st.session_state['_KR올림안내'] = (False, '씨티케이커넥트 분개장이나 연결패키지가 아닙니다 — 더존 분개장 · 「…연결패키지_씨티케이커넥트_…」 를 올려 주세요.')


def 자료받기(파일들):
    """올려 주신 파일을 안을 보고 제자리에 넣습니다 → (실적 엑셀 파일, [(이름, 종류)])."""
    실적, 목록 = None, []
    # 씨티케이커넥트 분개장 + 패키지를 함께 올리면 패키지를 양식으로 분개장을 반영 (2026-11-03)
    _kr패키지 = [f for f in 파일들 or [] if _KR패키지인가(f.getvalue()) and 한국법인자료(f.getvalue(), f.name)]
    _kr분개장 = [f for f in 파일들 or [] if 한국분개장인지(f.getvalue())]
    if _kr분개장:
        KR올린것처리(_kr패키지 + _kr분개장)
        for f in _kr패키지:
            목록.append((f.name, '한국실적'))
        for f in _kr분개장:
            _자리에쓰기(원장보관, f.getvalue(), f.name)
            목록.append((f.name, '한국분개장'))
        파일들 = [f for f in 파일들 if f not in _kr패키지 and f not in _kr분개장]
    _us원장 = [f for f in 파일들 or [] if _US원장인가(f.getvalue())]
    _us패키지 = [f for f in 파일들 or [] if _KR패키지인가(f.getvalue()) and not 한국법인자료(f.getvalue(), f.name)]
    if _us원장 and _us패키지:          # CTKonnect 원장 + 연결패키지를 함께 → 새 패키지 (2026-11-03)
        US올린것처리(_us패키지 + _us원장)
        for f in _us패키지:
            목록.append((f.name, '실적'))
        for f in _us원장:
            목록.append((f.name, '원장반영'))
        파일들 = [f for f in 파일들 if f not in _us원장 and f not in _us패키지]
    for f in 파일들 or []:
        바이트 = f.getvalue()
        종 = 자료종류(바이트)
        if 종 == '실적' and 한국법인자료(바이트, f.name):      # 2026-11-03 : 씨티케이커넥트 패키지는 KR 자리에
            종 = '한국실적'
            KR보관하기(바이트, f.name)
        elif 종 == '실적':
            실적 = f
            _자리에쓰기(전월보관, 바이트, f.name)
        elif 종 == '원장' and 한국분개장인지(바이트):      # 씨티케이커넥트 더존 분개장 → 패키지로 바꿔 KR 자리에
            _자리에쓰기(원장보관, 바이트, f.name)
            try:
                _됨, _글 = KR분개장반영(바이트, f.name)
            except Exception as e:
                _됨, _글 = False, f'분개장으로 패키지를 만들지 못했습니다 — {e}'
            st.session_state['_KR올림안내'] = (_됨, _글)
            종 = '한국분개장'
        elif 종 == '원장':
            _자리에쓰기(원장보관, 바이트, f.name)
        elif 종 == '양식':
            _자리에쓰기(양식보관, 바이트, f.name)
        목록.append((f.name, 종))
    return 실적, 목록


def 쓰던파일이름():
    """지금 화면이 쓰고 있는 실적 엑셀의 원래 파일 이름 (마지막으로 올린 것)."""
    try:
        with open(이름기록, encoding='utf-8') as f:
            return f.read().strip()
    except OSError:
        return ''

# 설정은 평소엔 접어 두고, 필요할 때만 펴서 씁니다 (사이드바를 깔끔하게).
# ※ 올려 둔 원장이 없으면(웹사이트가 다시 켜지면 지워집니다) 저절로 펼쳐서 바로 올릴 수 있게 합니다.
_용량 = 보관용량()
if _용량 is not None and _용량 >= 보관용량경고MB:
    st.sidebar.warning(f'자료 저장소(GitHub)가 {_용량:,.0f}MB 입니다 — 권장 1GB 에 가깝습니다. '
                       '클로드에게 「자료 저장소 정리」 를 요청해 주세요.')
설정칸 = st.sidebar.expander('⚙  설정 · 파일 업로드', expanded=not 자료있음)
설정칸.caption('화면 밝기')
설정칸.radio('화면 밝기', ['밝은 화면', '어두운 화면'], horizontal=True,
             key='ui_theme', label_visibility='collapsed')

설정칸.caption('올리는 곳 — 화면에 「마지막 올린 곳」으로 남습니다')
설정칸.radio('올리는 곳', ['씨티케이커넥트', 'CTKonnect'], key='올린이', index=None,
             horizontal=True, label_visibility='collapsed')
설정칸.caption('실적 엑셀 · 원장 · 패키지 양식을 한꺼번에 올리셔도 됩니다 — '
               '파일 안을 보고 알아서 제자리에 넣습니다.')
올린것들 = 설정칸.file_uploader('자료 올리기 (.xlsx · 여러 개 가능)', type=['xlsx'],
                                accept_multiple_files=True, key='설정업로드')
# 방금 올리신 것인지 확인해 둡니다 (이름+크기) — 새 자료일 때만 화면을 딱 한 번 새로 그립니다.
_올린도장 = tuple(sorted((str(f.name), int(getattr(f, 'size', 0) or 0))
                         for f in (올린것들 or [])))
_새로올림 = bool(올린것들) and st.session_state.get('_올린자료도장') != _올린도장
if not 올린것들:                       # 목록을 비우시면 다음에 같은 파일을 올려도 다시 읽습니다
    st.session_state.pop('_올린자료도장', None)

업로드파일, 가른것 = 자료받기(올린것들)

# 실적 엑셀은 어느 화면에서 올리셔도 곧바로 보관해 둡니다
# (실적보고·연결패키지 화면은 아래 본문까지 내려가지 않고 끝나기 때문입니다).
if 업로드파일 is not None:
    try:
        with open(CACHE_PATH, 'wb') as f:
            f.write(업로드파일.getvalue())
        with open(이름기록, 'w', encoding='utf-8') as f:
            f.write(업로드파일.name)
    except OSError:
        pass
    보관해두기(업로드파일.getvalue(), 업로드파일.name)

for _이름, _종 in 가른것:
    설정칸.caption(f'{"✔" if _종 else "?"} {_이름[:34]} → '
                   + 종류이름.get(_종, '무슨 자료인지 모르겠습니다'))

if _새로올림:
    st.session_state['_올린자료도장'] = _올린도장
    st.rerun()          # 올리자마자 새 자료로 화면을 다시 그립니다

# ※ 사이드바의 「이 대시보드 공유」 칸은 없앴습니다.
#   공유는 화면 오른쪽 아래 Streamlit 의 「Manage app」 → ⋮ → Settings → Sharing 에서 합니다.
#   그 단추는 웹사이트(streamlit.app)가 우리 화면 바깥에 직접 그리는 것이라
#   app.py 에서는 손댈 수 없습니다 — 브라우저에서 Streamlit 에 로그인해야 나타납니다.

# ══════════════════════════════════════════════════════════════
# 실적보고 엑셀작성 — 원장이 아직 없어도 열 수 있어야 하므로 여기서 처리합니다
# ══════════════════════════════════════════════════════════════
if 메뉴 == '실적보고 엑셀작성':
    st.html("""<div class="wrap" translate="no" style="margin-bottom:0">
  <div class="page-head"><span class="t">실적보고 엑셀작성</span>
    <span class="pill">원장 + 전월 자료 → 당월 연결패키지</span></div>
</div>""")
    # ══ CTKonnect (2026-10-29) — ① 퀵북 원장 · ② 통장(2026년 금융거래, 선택) · ③ 전월 연결패키지 → 결산월 연결패키지
    st.html(f"""<div class="wrap" translate="no">
  <div style="font-size:14.5px; line-height:1.9; color:{T['ink2']}; margin:2px 0 14px">
    <b style="color:{T['ink']}">①</b> 퀵북 원장(General Ledger) 과
    <b style="color:{T['ink']}">②</b> 통장 거래내역(「2026년 금융거래」 파일 · 선택) 을 올리고, 결산월 · 환율을 고른 뒤
    <b style="color:{T['ink']}">⚙ 만들기</b> 를 누르면 그 달 연결패키지가 만들어집니다.<br>
    <b style="color:{T['ink']}">③ 전월 연결패키지</b>(자료실의 가장 새 파일 — 사용자 마지막 버전) 를 그대로 두고
    「26년 원장_raw」 를 새 원장으로 갈아 끼우므로, BS · IS · 요약보고의 수식과 양식은 그대로 남고 엑셀이 열릴 때 다시 계산합니다.<br>
    통장 파일은 ① <b style="color:{T['ink']}">원장의 통장 거래와 맞는지 대사</b>(아래 표 — 올리면 바로 나옵니다) 와
    ② 「채권채무」 시트의 <b style="color:{T['ink']}">결산 후 회수 · 지급</b>(원장이 끝난 날 뒤 입출금)에 씁니다.
  </div></div>""")
    칸1, 칸2, 칸3 = st.columns(3, gap='small')
    with 칸1:
        st.caption('① 퀵북 원장 (General Ledger) 또는 씨티케이커넥트 더존 분개장 · 필수')
        _gl = st.file_uploader('원장', type=['xlsx'], key='ctk_gl', label_visibility='collapsed')
        if _gl is not None:
            _자리에쓰기(원장보관, _gl.getvalue(), _gl.name)
        _gl있음, _gl이름, _s1, _t1 = 보관자료(원장보관)
        st.caption(f'📁 {_gl이름}' if _gl있음 else '아직 없음')
    with 칸2:
        st.caption('② 통장 거래내역 (「2026년 금융거래」 파일 · 선택)')
        _bk = st.file_uploader('통장', type=['xlsx'], key='ctk_bank', label_visibility='collapsed')
        if _bk is not None:
            _자리에쓰기(통장보관, _bk.getvalue(), _bk.name)
        _bk있음, _bk이름, _s2, _t2 = 보관자료(통장보관)
        st.caption(f'📁 {_bk이름}' if _bk있음 else '없으면 원장 안의 입출금만 씁니다')
    with 칸3:
        st.caption('③ 전월 연결패키지 (자동 — 다른 파일을 쓰려면 올려 주세요)')
        _pv = st.file_uploader('전월', type=['xlsx'], key='ctk_prev', label_visibility='collapsed')
        if _pv is not None:
            _전월바이트, _전월이름 = _pv.getvalue(), _pv.name
        else:
            _최신 = _자료실최신패키지()
            if _최신:
                with open(_최신, 'rb') as _f:
                    _전월바이트 = _f.read()
                _전월이름 = os.path.basename(_최신)
            elif os.path.exists(CACHE_PATH):
                with open(CACHE_PATH, 'rb') as _f:
                    _전월바이트 = _f.read()
                _전월이름 = 쓰던파일이름() or CACHE_PATH
            else:
                _전월바이트, _전월이름 = None, ''
        st.caption(f'📁 {_전월이름}' if _전월바이트 else '자료실에 연결패키지가 없습니다 — 올려 주세요')

    # ── 씨티케이커넥트 : ① 에 더존 분개장을 올리면 씨티케이커넥트 연결패키지를 만듭니다 (2026-11-03)
    if _gl있음 and 한국분개장인지(보관바이트(원장보관)):
        if _pv is not None and 한국법인자료(_pv.getvalue(), _pv.name):
            _KR전월, _KR전월이름 = _pv.getvalue(), _pv.name
        else:
            _KR전월, _KR전월이름 = _KR패키지바이트()
        st.html(f"""<div class="wrap" translate="no"><div class="card" style="margin-top:4px">
          <h3>씨티케이커넥트 연결패키지 만들기 <span class="unitbadge">단위 원</span></h3>
          <div class="sub">① 더존 분개장 「{_html.escape(str(_gl이름))}」 + ③ 전월 씨티케이커넥트 패키지
            「{_html.escape(str(_KR전월이름 or '없음'))}」 → 「26년 원장_raw」 를 새 분개장으로 갈아 끼웁니다.
            BS · IS · 요약보고 수식과 양식, BS_IS_매핑, 활동분류(같은 전표 줄)는 전월 패키지 것을 그대로 씁니다. 통장 파일 · 환율은 쓰지 않습니다.</div>
        </div></div>""")
        if not _KR전월:
            st.info('③ 에 씨티케이커넥트 연결패키지(「…연결패키지_씨티케이커넥트_…」)를 올려 주세요.')
            st.stop()
        try:
            _kd, _ = _KR분개장읽기(보관바이트(원장보관))
        except Exception as e:
            st.error(f'분개장을 읽지 못했습니다 — {e}')
            st.stop()
        _끝 = _kd['일자'].max()
        _연, _결산월 = int(_끝.year), int(_끝.month)
        _있는번호 = [int(m.group(1)) for f in glob.glob(os.path.join(자료실, '**', '*.xlsx'), recursive=True)
                     for m in [re.search(rf'_연결패키지_씨티케이커넥트_{_결산월}월말_v(\d+)', os.path.basename(f))] if m]
        _m2 = re.search(rf'_연결패키지_씨티케이커넥트_{_결산월}월말_v(\d+)', str(_KR전월이름 or ''))
        if _m2:
            _있는번호.append(int(_m2.group(1)))
        _버1, _버2 = st.columns([1, 3.2], gap='small')
        _버전 = _버1.number_input(f'{_결산월}월말 버전 번호', min_value=1, step=1, value=max(_있는번호 or [0]) + 1,
                                  key=f'kr_ver_{_결산월}', help='기본값 = 같은 달 씨티케이커넥트 패키지의 마지막 번호 + 1')
        _버2.caption(f'분개장 {_kd["일자"].min():%Y-%m-%d} ~ {_끝:%Y-%m-%d} · {len(_kd):,}줄 · 결산월 {_결산월}월 · '
                     f'파일 이름 : 「…연결패키지_씨티케이커넥트_{_결산월}월말_v{int(_버전)}.xlsx」')
        if st.button('⚙  씨티케이커넥트 연결패키지 만들기', type='primary', key='kr_make'):
            try:
                with st.spinner('씨티케이커넥트 연결패키지를 만드는 중입니다 (20~40초)...'):
                    _결과, _안내 = KR패키지만들기(보관바이트(원장보관), _KR전월)
            except Exception as e:
                st.error(f'만들지 못했습니다 — {e}')
                st.stop()
            _q = (_결산월 - 1) // 3 + 1
            _폴더 = os.path.join(자료실, f'{_연}년 {_결산월}월', '씨티케이커넥트')
            _n = int(_버전)
            _이름 = f'{_연}. {_q}Q_연결패키지_씨티케이커넥트_{_결산월}월말_v{_n}.xlsx'
            while os.path.exists(os.path.join(_폴더, _이름)):        # 같은 이름이 있으면 덮어쓰지 않고 다음 번호
                _n += 1
                _이름 = f'{_연}. {_q}Q_연결패키지_씨티케이커넥트_{_결산월}월말_v{_n}.xlsx'
            try:
                os.makedirs(_폴더, exist_ok=True)
                with open(os.path.join(_폴더, _이름), 'wb') as _f:
                    _f.write(_결과)
                _안내['저장'] = os.path.join(_폴더, _이름)
            except OSError as e:
                _안내['저장'] = f'저장 못 함 ({e})'
            KR보관하기(_결과, _이름)          # 씨티케이커넥트 화면 · 웹 보관소에도 바로 반영
            st.session_state['_KR결과'] = (_결과, _이름, _안내)
        if st.session_state.get('_KR결과'):
            _결과, _이름, _안내 = st.session_state['_KR결과']
            _알림 = ''.join(f'<li>{_html.escape(a)}</li>' for a in _안내.get('알림', []))
            _새 = ''.join(f'<li style="color:#C00000">새 계정 {_html.escape(a)} — BS_IS_매핑 확인 필요</li>' for a in _안내.get('새계정', []))
            st.html(f"""<div class="wrap" translate="no"><div class="card" style="margin-top:4px">
              <h3>만들었습니다 — {_html.escape(_이름)}</h3>
              <div class="sub">분개장 {_안내['줄']:,}줄 · 전월에서 이어받은 활동분류 {_안내.get('이어받음', 0):,}줄 ·
                저장 : {_html.escape(str(_안내.get('저장', '')))} · 엑셀에서 한 번 열어 저장하시면 수식 값이 채워집니다.</div>
              <ul style="margin:8px 0 0; padding-left:18px">{_새}{_알림}</ul></div></div>""")
            st.download_button('⤓  씨티케이커넥트 연결패키지 내려받기', data=_결과, file_name=_이름,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='kr_down')
        st.stop()

    if not _gl있음 or not _전월바이트:
        st.info('① 원장과 ③ 전월 연결패키지가 있어야 만들 수 있습니다.')
        st.stop()
    try:
        _g = 커_원장읽기(보관바이트(원장보관))
    except Exception as e:
        st.error(f'원장을 읽지 못했습니다 — {e}')
        st.stop()
    # ── 통장 ↔ 원장 대사 (통장 파일이 있을 때) ─────────────────────────────
    if _bk있음:
        try:
            _대사 = 커_통장대사(_g, 커_통장읽기(보관바이트(통장보관)))
        except Exception as e:
            _대사 = None
            st.warning(f'통장 대사를 하지 못했습니다 — {e}')
        if _대사:
            _줄 = ''
            _안 = ''
            for r in _대사:
                차이 = len(r['통장만']) + len(r['원장만'])
                판정 = ('<span style="color:#2e7d32;font-weight:800">✔ 일치</span>' if 차이 == 0 else
                        f'<span style="color:{T["warn"]};font-weight:800">⚠ {차이}건 다름</span>')
                _줄 += (f'<tr class="sub"><td class="lft">{_html.escape(r["계좌"])}</td>'
                        f'<td>{r["통장"]["건"]:,}</td><td>{금액(r["통장"]["입"])}</td><td>{금액(r["통장"]["출"])}</td>'
                        f'<td>{r["원장"]["건"]:,}</td><td>{금액(r["원장"]["입"])}</td><td>{금액(r["원장"]["출"])}</td>'
                        f'<td>{금액(r["통장"]["입"] - r["통장"]["출"] - r["원장"]["입"] + r["원장"]["출"])}</td>'
                        f'<td style="text-align:center;background:transparent">{판정}</td></tr>')
                for 종류, 목록 in (('통장에만 있음', r['통장만']), ('원장에만 있음', r['원장만'])):
                    for x in 목록:
                        _안 += (f'<tr class="sub"><td class="lft">{_html.escape(r["계좌"][:22])}</td><td>{종류}</td>'
                                f'<td>{x["일자"]:%Y-%m-%d}</td><td>{금액(x["금액"])}</td>'
                                f'<td class="lft" style="font-weight:400;background:transparent">{_html.escape(str(x["적요"])[:80])}</td></tr>')
            _끝 = _대사[0]['기간끝']
            st.html(f"""<div class="wrap" translate="no"><div class="card" style="margin-top:4px">
              <h3>통장 ↔ 원장 대사 <span class="unitbadge">단위 USD</span></h3>
              <div class="sub">원장 보통예금 계정 거래 ↔ 통장(「2026년 금융거래」 자금data(US)) 거래 · {_끝:%Y-%m-%d} 까지 (두 자료가 다 있는 날까지) ·
                같은 금액이 3일 안에 있으면 짝, 남은 것은 같은 날 합계로 다시 맞춤 · 금액 0 줄은 뺌</div>
              <div class="scrollx"><table class="lined" style="width:100%; table-layout:fixed;">
                <colgroup><col style="width:22%"><col style="width:7%"><col style="width:11%"><col style="width:11%"><col style="width:7%">
                  <col style="width:11%"><col style="width:11%"><col style="width:9%"><col style="width:11%"></colgroup>
                <thead><tr><th>계좌</th><th>통장 건수</th><th>통장 입금</th><th>통장 출금</th><th>원장 건수</th><th>원장 입금</th>
                  <th>원장 출금</th><th>순액 차이</th><th>판정</th></tr></thead>
                <tbody>{_줄}</tbody></table></div>
              {(f'<div class="scrollx" style="margin-top:12px"><table class="lined" style="width:100%; table-layout:fixed;">'
                f'<colgroup><col style="width:20%"><col style="width:12%"><col style="width:11%"><col style="width:12%"><col style="width:45%"></colgroup>'
                f'<thead><tr><th>계좌</th><th>구분</th><th>일자</th><th>금액</th><th>적요</th></tr></thead><tbody>{_안}</tbody></table></div>')
               if _안 else '<div class="sub" style="margin-top:8px">건별로도 모두 짝이 맞았습니다 (통장 여러 줄 = 원장 한 줄처럼 같은 날 합계로 맞춘 것 포함).</div>'}
            </div></div>""")

    _달들 = sorted(_g['일자'].dt.month.unique().tolist())
    # 기본 결산월 : 파일 이름의 「○월결산」 → 없으면 매출(511…) 이 잡힌 마지막 달
    _m = re.search(r'(\d{1,2})\s*월\s*결산', str(_gl이름))
    _매출 = _g[_g['계정'].str.startswith('511')]['일자']
    _기본달 = int(_m.group(1)) if _m else int((_매출 if len(_매출) else _g['일자']).max().month)
    선1, 선2, 선3, 선4 = st.columns([1, 1, 1, 1.2], gap='small')
    _결산월 = 선1.selectbox('결산월', _달들, index=_달들.index(_기본달) if _기본달 in _달들 else len(_달들) - 1,
                            format_func=lambda m: f'{m}월', key='ctk_month',
                            help=f'원장 마지막 날 {_g["일자"].max():%Y-%m-%d} · 기본값 = 원장 파일 이름의 「○월결산」(없으면 매출이 잡힌 마지막 달)')
    _월말 = 선2.number_input(f'{_결산월}월말 환율 (원/USD)', min_value=0.0, value=0.0, step=0.01, format='%.2f', key='ctk_fx1',
                             help='서울외국환중개 매매기준율 — 0 이면 전월 값 그대로 둡니다')
    _평균 = 선3.number_input(f'1~{_결산월}월 평균환율 (원/USD)', min_value=0.0, value=0.0, step=0.01, format='%.2f', key='ctk_fx2',
                             help='0 이면 전월 값 그대로 둡니다')
    with 선4:
        st.caption('올리는 곳')
        st.radio('올리는 곳', ['씨티케이커넥트', 'CTKonnect'], key='올린곳_작성', index=None,
                 horizontal=True, label_visibility='collapsed')

    _있는번호 = [int(m.group(1)) for f in glob.glob(os.path.join(자료실, '**', '*.xlsx'), recursive=True)
                 for m in [re.search(rf'_연결패키지_CTKonnect_{_결산월}월말_v(\d+)', os.path.basename(f))] if m]
    for _올린이름 in [str(_전월이름 or ''), 쓰던파일이름() or '']:
        _m2 = re.search(rf'_연결패키지_CTKonnect_{_결산월}월말_v(\d+)', _올린이름)
        if _m2:
            _있는번호.append(int(_m2.group(1)))
    _버1, _버2 = st.columns([1, 3.2], gap='small')
    _버전 = _버1.number_input(f'{_결산월}월말 버전 번호', min_value=1, step=1, value=max(_있는번호 or [0]) + 1,
                              key=f'ctk_ver_{_결산월}',
                              help='기본값 = 자료실에 있는 같은 달 연결패키지의 마지막 번호 + 1')
    _버2.caption(f'파일 이름 : 「…연결패키지_CTKonnect_{_결산월}월말_v{int(_버전)}.xlsx」 · '
                 + (f'자료실 마지막 = v{max(_있는번호)}' if _있는번호 else
                    '자료실에서 같은 달 파일을 못 찾았습니다 (웹사이트에서는 PC 자료실 마지막 번호 + 1 을 직접 넣어 주세요)'))

    if st.button('⚙  연결패키지 만들기', type='primary', key='ctk_make'):
        try:
            with st.spinner('연결패키지를 만드는 중입니다 (10~20초)...'):
                _결과, _안내 = 커_패키지만들기(보관바이트(원장보관), _전월바이트, int(_결산월),
                                             _월말 or None, _평균 or None,
                                             보관바이트(통장보관) if _bk있음 else None)
                try:
                    _결과 = 과제잠그기(_결과, (int(_안내['연도']), int(_결산월)))
                except Exception:
                    pass
            _연 = int(_안내['연도'])
            _폴더 = os.path.join(자료실, f'{_연}년 {_결산월}월', 'CTKonnect')     # 2026-11-03 : 달 폴더 안 법인별 하위 폴더
            _이름 = f'{_연}. {(_결산월 - 1) // 3 + 1}Q_연결패키지_CTKonnect_{_결산월}월말_v{int(_버전)}.xlsx'
            if os.path.exists(os.path.join(_폴더, _이름)):      # 같은 이름이 이미 있으면 덮어쓰지 않고 다음 번호로
                _n = int(_버전)
                while os.path.exists(os.path.join(_폴더, _이름)):
                    _n += 1
                    _이름 = f'{_연}. {(_결산월 - 1) // 3 + 1}Q_연결패키지_CTKonnect_{_결산월}월말_v{_n}.xlsx'
            try:
                os.makedirs(_폴더, exist_ok=True)
                with open(os.path.join(_폴더, _이름), 'wb') as _f:
                    _f.write(_결과)
                _안내['저장'] = os.path.join(_폴더, _이름)
            except OSError as e:
                _안내['저장'] = f'저장 못 함 ({e})'
            try:
                자료기록쓰기(_이름, str(st.session_state.get('올린곳_작성') or ''))
            except Exception:
                pass
            st.session_state['ctk_결과'] = (_결과, _이름, _안내)
        except Exception as e:
            st.error(f'만들지 못했습니다 — {e}')
    if st.session_state.get('ctk_결과'):
        _결과, _이름, _안내 = st.session_state['ctk_결과']
        _줄 = ''.join(f'<li>{_html.escape(str(x))}</li>' for x in _안내.get('알림', []))
        st.html(f"""<div class="wrap" translate="no"><div class="card" style="margin-top:8px">
          <h3>만든 연결패키지 <span class="unitbadge">{_html.escape(_이름)}</span></h3>
          <div class="sub">{_안내['연도']}년 {_안내['결산월']}월말 · 원장 {_안내['원장건수']:,}건 (원장 마지막 날 {_안내['원장끝']:%Y-%m-%d})
            · 활동분류를 새로 붙인 손익 거래 {_안내['새분류']:,}건 · 저장 : {_html.escape(str(_안내.get('저장', '')))}</div>
          <ul style="font-size:13.5px; line-height:1.8; color:{T['ink2']}">{_줄}</ul>
          <div class="sub">엑셀에서 한 번 열어 저장하시면 수식 값이 채워지고, 대시보드도 그 숫자로 그립니다 (자료실 가장 새 파일을 자동으로 읽음).</div>
        </div></div>""")
        st.download_button('⤓  연결패키지 내려받기', data=_결과, file_name=_이름,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='ctk_dl')
    st.stop()
    st.html(f"""<div class="wrap" translate="no">
  <div style="font-size:14.5px; line-height:1.9; color:{T['ink2']}; margin:2px 0 14px">
    <b style="color:{T['ink']}">①</b> QuickBooks 원장 원본과
    <b style="color:{T['ink']}">②</b> 지난달 연결패키지(또는 실적보고 엑셀)를 올리고
    <b style="color:{T['ink']}">⚙ 생성</b> 을 누르면 이번 달 자료가 만들어집니다.<br>
    올려주신 파일을 <b style="color:{T['ink']}">그대로 두고 「원장_raw」 시트만 새 원장으로 갈아끼우므로</b>,
    연결패키지의 BS·IS(원장을 보는 수식)·나머지 시트가 그대로 남고, 엑셀이 열릴 때 스스로 다시 계산합니다.<br>
    만든 뒤 <b style="color:{T['ink']}">✓ 적용하기</b> 를 누르면 대시보드 기준 자료가 되고,
    다음 달 ② 전월 자료로도 기억해 두어 다시 올리지 않아도 됩니다.
  </div>
</div>""")

    # ② 칸은 「적용하기」 뒤에 비워야 하므로 세대 번호를 붙인 열쇠를 씁니다
    전월칸열쇠 = f"gen_prev_{st.session_state.get('전월칸세대', 0)}"
    작성좌, 작성우 = st.columns(2, gap='medium')
    with 작성좌:
        st.markdown('**① 원장 원본** (QuickBooks General Ledger)')
        st.caption('퀵북에서 1월~당월 누적 General Ledger List 를 내려받은 엑셀 — '
                   '**같은 파일에 당월 퀵북 BS·IS 시트(「8월(BS)」「8월(IS)」)도 함께** 넣어 주세요 '
                   '(연결재무제표 패키지 화면에서 씁니다)')
        원장올림 = st.file_uploader('원장 원본', type=['xlsx'], key='gen_gl',
                                    label_visibility='collapsed')
    with 작성우:
        st.markdown('**② 전월 자료** (지난달 연결패키지)')
        st.caption('지난달에 만든 「…연결패키지_CTKonnect_○월말_v○.xlsx」 — '
                   '지난달 「✓ 적용하기」 를 눌렀다면 이미 기억하고 있어 안 올려도 됩니다 '
                   '(처음엔 실적보고 엑셀도 됩니다)')
        전월올림 = st.file_uploader('전월 자료', type=['xlsx'], key=전월칸열쇠,
                                    label_visibility='collapsed')
    # 올려 주신 파일은 그대로 남겨 두어, 다음에는 다시 올리지 않아도 됩니다
    자료남기기(원장보관, 원장올림)
    자료남기기(전월보관, 전월올림)
    원장있음, 원장이름, 원장크기, _t1 = 보관자료(원장보관)
    전월있음, 전월이름, 전월크기, _t2 = 보관자료(전월보관)

    보관목록 = 자료실목록()
    패키지들 = 패키지목록()
    # 「적용하기」 기본 대상 — 실적보고 엑셀과 연결패키지 가운데 가장 최근 것
    최신 = max(보관목록[:1] + 패키지들[:1], key=lambda d: d['시각'], default=None)
    둘다 = 원장있음 and 전월있음

    # ── 어떤 자료로 만드는지 눈으로 확인할 수 있게 보여 줍니다
    def _올린줄(번호, 파일, 설명, 있음, 이름, 크기):
        if 파일 is None and not 있음:
            return (f'<tr class="sub"><td class="name">{번호} {설명}</td>'
                    f'<td style="color:{T["ink3"]}">아직 올리지 않음</td></tr>')
        새로 = 파일 is not None
        이름 = 파일.name if 새로 else 이름
        크기 = 파일.size if 새로 else 크기
        표 = ('' if 새로 else
              f'<span style="color:{T["accent"]}; font-weight:700">  ·  지난번에 올린 자료</span>')
        return (f'<tr class="sub"><td class="name">{번호} {설명}</td>'
                f'<td class="lft" style="font-weight:700">{이름}'
                f'<span style="color:{T["ink3"]}; font-weight:400">'
                f'  ·  {크기/1e6:,.1f}MB</span>{표}</td></tr>')

    st.html(f"""<div class="wrap" translate="no">
  <div class="card" style="margin-top:6px">
    <h3>이 자료로 만듭니다</h3>
    <div class="sub">한 번 올린 파일은 그대로 기억합니다 — 바뀐 것만 새로 올리시면 됩니다</div>
    <table class="lined" style="margin-top:8px">
      <colgroup><col style="width:32%"><col style="width:68%"></colgroup>
      <tbody>
        {_올린줄('①', 원장올림, '원장 원본', 원장있음, 원장이름, 원장크기)}
        {_올린줄('②', 전월올림, '전월 자료 (연결패키지)', 전월있음, 전월이름, 전월크기)}
      </tbody>
    </table>
  </div>
</div>""")

    # ── 기억해 둔 자료는 언제든 다시 내려받을 수 있게 합니다
    보관줄 = [('② 전월 자료', 전월보관), ('① 원장 원본', 원장보관),
              ('연결패키지 양식', 양식보관)]
    있는것 = [(라벨, 자리) for 라벨, 자리 in 보관줄 if 보관자료(자리)[0]]
    if 있는것:
        st.caption('보관 중인 자료 내려받기 — 언제 올리셨든 마지막 것을 그대로 드립니다')
        받는칸 = st.columns(len(있는것) + max(0, 3 - len(있는것)), gap='small')
        for 칸, (라벨, 자리) in zip(받는칸, 있는것):
            _있, 이름, 크기, 때 = 보관자료(자리)
            바이트 = 보관바이트(자리)
            if 바이트 is None:
                continue
            칸.download_button(
                f'⤓  {라벨}', data=바이트, file_name=이름 or os.path.basename(자리),
                width='stretch', key='보관받기_' + 자리,
                mime=('application/vnd.openxmlformats-officedocument'
                      '.spreadsheetml.sheet'),
                help=f'{이름}  ·  {크기/1e6:,.1f}MB  ·  '
                     f'{datetime.datetime.fromtimestamp(때):%Y.%m.%d %H:%M} 에 올림')

    # ── 생성 · 내려받기
    def _생성누름():
        st.session_state['엑셀생성요청'] = True

    만들칸, 받을칸, 적용칸, 설명칸2 = st.columns([1, 1, 1.3, 1.6], gap='small')
    with 만들칸:
        st.button('⚙  생성', width='stretch', type='primary', key='보고서생성',
                  disabled=not 둘다, on_click=_생성누름,
                  help='위 「이 자료로 만듭니다」 의 두 파일로 오늘 날짜의 다음 번호를 붙여 만듭니다')

    만든것 = None
    if st.session_state.pop('엑셀생성요청', False) and 둘다:
        try:
            만든바이트, 만든안내 = 실적보고엑셀캐시(
                원장올림.getvalue() if 원장올림 is not None else 보관바이트(원장보관),
                전월올림.getvalue() if 전월올림 is not None else 보관바이트(전월보관))
            if 만든안내.get('패키지형'):        # 전월 자료가 연결패키지면 결과도 연결패키지 이름으로
                버전 = 패키지다음버전(만든안내['월끝'])
                이름 = 패키지파일이름(만든안내['연도'], 만든안내['월끝'], 버전)
            else:
                버전 = 다음버전()
                이름 = 보고서파일이름(만든안내['연도'], 만든안내['월끝'], 버전)
            저장됨 = ''
            try:
                os.makedirs(자료실, exist_ok=True)
                with open(os.path.join(자료실, 이름), 'wb') as f:
                    f.write(만든바이트)
                저장됨 = os.path.abspath(os.path.join(자료실, 이름))
            except OSError as e:
                st.session_state['엑셀저장오류'] = str(e)
            만든것 = (이름, 만든바이트, 만든안내, 버전, 저장됨)
            st.session_state['엑셀마지막안내'] = 만든안내
            # 화면을 새로 그려도 「이 자료로 대시보드 바꾸기」 단추가 이 파일을 가리키도록
            st.session_state['방금만든자료'] = {'이름': 이름, '경로': 저장됨}
            보관목록 = 자료실목록()
            패키지들 = 패키지목록()
            최신 = max(보관목록[:1] + 패키지들[:1], key=lambda d: d['시각'], default=None)
        except Exception as e:
            st.session_state['엑셀생성오류'] = str(e)

    오류 = st.session_state.pop('엑셀생성오류', None)
    if 오류:
        st.html(f"""<div class="wrap" translate="no">
  <div class="card" style="max-width:860px; margin-top:12px; border-left:4px solid {ROSE}">
    <h3>만들지 못했습니다</h3>
    <div style="font-size:14px; line-height:1.9; color:{T['ink2']}; margin-top:8px">{오류}</div>
  </div>
</div>""")

    with 받을칸:
        if 만든것 is not None:
            st.download_button(f'⤓  {오늘날짜():%Y.%m.%d}_v{만든것[3]} 내려받기',
                               data=만든것[1], file_name=만든것[0],
                               mime=('application/vnd.openxmlformats-officedocument'
                                     '.spreadsheetml.sheet'), width='stretch')
        elif 최신 is not None:
            try:
                날 = 최신.get('만든날')
                이름표 = (f"{날:%Y.%m.%d}_v{최신['버전']}" if 날 else f"v{최신['버전']}")
                with open(최신['경로'], 'rb') as f:
                    st.download_button(f'⤓  {이름표} 내려받기', data=f.read(),
                                       file_name=최신['이름'],
                                       mime=('application/vnd.openxmlformats-officedocument'
                                             '.spreadsheetml.sheet'), width='stretch')
            except OSError:
                st.button('⤓  내려받기', width='stretch', disabled=True)
        elif 전월있음:
            # 자료실이 비어 있어도(웹사이트가 다시 켜진 뒤) 기억해 둔 마지막 자료는 드립니다
            바이트 = 보관바이트(전월보관)
            if 바이트:
                st.download_button('⤓  마지막 자료 내려받기', data=바이트,
                                   file_name=전월이름 or '실적보고.xlsx', width='stretch',
                                   key='마지막자료받기',
                                   mime=('application/vnd.openxmlformats-officedocument'
                                         '.spreadsheetml.sheet'),
                                   help=f'기억해 둔 「{전월이름}」 입니다')
            else:
                st.button('⤓  내려받기', width='stretch', disabled=True)
        else:
            st.button('⤓  내려받기', width='stretch', disabled=True,
                      help='아직 만든 자료가 없습니다. 두 파일을 올리고 「생성」을 눌러 주세요')
    # ── 만든 자료를 곧바로 대시보드 기준 자료로 삼습니다 (사이드바에 다시 올릴 필요 없음)
    방금 = st.session_state.get('방금만든자료') or {}
    적용경로 = (만든것[4] if 만든것 is not None else '') or 방금.get('경로') \
                or (최신 or {}).get('경로') or ''
    적용이름 = (만든것[0] if 만든것 is not None else '') or 방금.get('이름') \
                or (최신 or {}).get('이름') or ''
    with 적용칸:
        st.radio('올리는 곳', ['씨티케이커넥트', 'CTKonnect'], key='올린곳_작성', index=None,
                 horizontal=True, label_visibility='collapsed',
                 help='누가 적용했는지 사이드바 「마지막 올린 곳」에 남습니다')
        if 적용경로 and os.path.exists(적용경로):
            if st.button('✓  적용하기', width='stretch',
                         type='primary', key='대시보드적용',
                         help=f'「{적용이름}」 을 대시보드 기준 자료로 삼고, 다음 달 ② 전월 자료로도 '
                              f'기억해 둡니다 — 다시 올리지 않으셔도 됩니다'):
                try:
                    with open(적용경로, 'rb') as f:
                        새바이트 = f.read()
                    with open(CACHE_PATH, 'wb') as f:
                        f.write(새바이트)
                    try:
                        with open(이름기록, 'w', encoding='utf-8') as f:
                            f.write(적용이름)
                    except OSError:
                        pass
                    # 다음 달 ② 전월 자료로도 기억해 둡니다 (② 칸은 비워서 이 파일이 쓰이게)
                    try:
                        with open(전월보관, 'wb') as f:
                            f.write(새바이트)
                        with open(전월보관 + '.txt', 'w', encoding='utf-8') as f:
                            f.write(적용이름)
                        st.session_state.pop('보관도장_' + 전월보관, None)
                        st.session_state['전월칸세대'] = st.session_state.get('전월칸세대', 0) + 1
                    except OSError:
                        pass
                    보관해두기(새바이트, 적용이름)   # 웹 보관소 · 「마지막 올린 곳 · 때」 기록
                    st.cache_data.clear()          # 새 자료로 다시 읽도록 비웁니다
                    st.session_state['대시보드바꿈'] = 적용이름
                    st.rerun()
                except OSError as e:
                    st.session_state['적용오류'] = str(e)
        else:
            st.button('✓  적용하기', width='stretch', disabled=True,
                      key='대시보드적용막음',
                      help='먼저 「⚙ 생성」 으로 이번 달 자료를 만들어 주세요')
    with 설명칸2:
        if 최신 is not None or 만든것 is not None or 둘다:
            st.caption('「✓ 적용하기」를 누르면 대시보드가 이번 달 숫자로 바뀌고, '
                       '이 파일이 다음 달 ② 전월 자료가 됩니다 (다시 올리지 않아도 됨).')
        else:
            st.caption('두 파일을 모두 올리면 「생성」 단추가 켜집니다.')
        st.caption('파일 이름 — ② 가 연결패키지면 「2026. 3Q_연결패키지_CTKonnect_8월말_v번호」'
                   '(같은 달 안에서 v1·v2…), 실적보고 엑셀이면 '
                   f'「…클로드작성_{오늘날짜():%Y.%m.%d}_v번호」 입니다.')

    if 적용이름 and os.path.exists(적용경로):
        st.caption(f'대시보드 바꾸기에 쓸 자료 — {적용이름}')
    바꿈 = st.session_state.pop('대시보드바꿈', None)
    if 바꿈:
        st.success(f'대시보드 기준 자료를 「{바꿈}」 로 바꿨습니다. '
                   f'왼쪽에서 「실적보고」를 누르시면 새 숫자로 보입니다.')
    적용오류 = st.session_state.pop('적용오류', None)
    if 적용오류:
        st.caption(f'※ 대시보드 자료를 바꾸지 못했습니다({적용오류}).')

    저장오류 = st.session_state.pop('엑셀저장오류', None)
    if 저장오류:
        st.caption(f'※ 자료실 폴더에 저장하지 못했습니다({저장오류}). '
                   f'위 내려받기 단추로 받아 주세요.')

    # ── 방금 만든 결과 요약
    if 만든것 is not None:
        이름, _b, 안내, 버전, 저장됨 = 만든것
        st.html(f"""<div class="wrap" translate="no">
  <div class="card" style="margin-top:14px">
    <h3>v{버전} 파일을 만들었습니다</h3>
    <div class="sub">{이름}</div>
    <table class="lined" style="margin-top:10px">
      <colgroup><col style="width:46%"><col style="width:54%"></colgroup>
      <tbody>
        <tr class="sub"><td class="name">대상 기간</td><td class="lft">{안내['기간']}</td></tr>
        <tr class="sub"><td class="name">갈아끼운 시트</td>
          <td class="lft">{안내['올해시트']} · {안내['행수']:,}줄 (거래 {안내['거래건수']:,}건)</td></tr>
        <tr class="sub"><td class="name">그대로 옮긴 시트</td>
          <td class="lft">{안내['시트수'] - 1}개 (피벗·차트 포함)</td></tr>
        <tr class="sub"><td class="name">보고 기준 월</td>
          <td class="lft">{안내['보고달']}월
            {'— 「' + str(안내['뺀시트']) + '」 시트는 이제 쓰지 않으므로 '
             '결과 파일에서 뺐습니다' if 안내.get('뺀시트')
             else '— 원장 마지막 거래일 기준'}</td></tr>
        {(f'<tr class="sub"><td class="name">연결패키지 BS · IS</td>'
          f'<td class="lft">당기 열은 「{안내["올해시트"]}」 를 보는 SUMIFS 수식 — '
          f'제목·환율 라벨을 {안내["보고달"]}월로 바꿨고, 숫자는 엑셀에서 열 때 다시 계산됩니다. '
          f'환율(M5)·수정분개(J·K)·CE·MC·CF 는 전월 값 그대로이니 확인해 주세요</td></tr>')
         if 안내.get('패키지형') else ''}
        {(f'<tr class="sub"><td class="name">퀵북 BS · IS → 「검증」</td>'
          f'<td class="lft">{안내["퀵북"]} — 「검증」 시트(BS 앞)에서 패키지와 일치 여부를 보세요</td></tr>')
         if 안내.get('퀵북') else ''}
        <tr class="sub"><td class="name">26년 자금 시트</td>
          <td class="lft">{(
            f"실제 현금흐름 {안내['자금']['현금줄']:,}줄 "
            f"(기초 {안내['자금']['기초']:,.2f} → {안내['보고달']}월말 "
            f"{안내['자금']['기말']:,.2f}) · 예상은 "
            + (f"{안내['자금']['예상월'][0]}~12월 "
               f"{안내['자금']['항목수']}개 항목"
               if 안내['자금'].get('예상월') else '없음')
            ) if 안내['자금'].get('현금줄') is not None
            else ('전월 자료에 없는 시트라 건너뜀' if 안내['자금'].get('없음')
                  else '다시 쓰지 못했습니다 — ' + str(안내['자금'].get('오류', '시트를 찾지 못함')))}
          </td></tr>
        <tr class="sub"><td class="name">전월에서 이어받은 활동분류</td>
          <td class="lft">{안내['이어받음']:,}건</td></tr>
        <tr class="sub"><td class="name">새 거래를 스스로 나눈 것</td>
          <td class="lft">{안내['자동분류']:,}건
            <span style="color:{T['ink3']}">
              (적요 {안내['분류내역'].get('적요일치', 0):,} ·
               거래처 {안내['분류내역'].get('거래처일치', 0):,} ·
               계정 {안내['분류내역'].get('계정과목대표값', 0):,} ·
               계정 최빈값 {안내['분류내역'].get('계정과목최빈값', 0):,})</span></td></tr>
        <tr class="sub"><td class="name">사람이 꼭 봐야 할 줄</td>
          <td class="lft" style="{'font-weight:800; color:' + ROSE
                                  if 안내['확인필요'] else ''}">{안내['확인필요']:,}건
            <span style="color:{T['ink3']}">
              (여러 갈래로 나뉘던 계정 {안내['분류내역'].get('확인필요', 0):,} ·
               처음 보는 계정 {안내['분류내역'].get('신규계정', 0):,})
              — T열에서 「확인필요」로 걸러 보세요</span></td></tr>
        <tr class="sub"><td class="name">전월 자료에 없던 새 거래</td>
          <td class="lft">{안내['새거래']:,}건</td></tr>
        <tr class="total"><td class="name">저장 위치</td>
          <td class="lft">{저장됨 or '(저장 안 됨 — 내려받기로 받아 주세요)'}</td></tr>
      </tbody>
    </table>
    <div style="margin-top:14px; padding-top:12px; border-top:1px solid {T['line']};
                font-size:13px; line-height:1.85; color:{T['ink3']}">
      원장에서 저절로 계산되는 것 — 금액 · 년 · 월 · 계정분류 · 분류 · 계정과목 · 보고금액 · 월(숫자)
      <b style="color:{T['ink2']}">(K~V열 수식 그대로)</b><br>
      전월에서 이어받는 것 — 활동분류 · 활동세부 · 정부지원 · 검토메모<br>
      새 거래는 지난달까지 나눠 두신 기준을 배워서 채우고, 그 근거를
      <b style="color:{T['ink2']}">T열 「클로드 검토」</b> 에 적어 둡니다
      (「확인필요」가 붙은 줄부터 봐 주세요)
    </div>
  </div>
</div>""")
        경고 = []
        if 안내['확인필요']:
            경고.append(f"활동분류를 확실하게 정하지 못한 줄이 {안내['확인필요']:,}건 있습니다. "
                        f"엑셀 「{안내['올해시트']}」 시트에서 T열을 「확인필요」로 걸러 보시면 "
                        f"바로 찾으실 수 있습니다.")
        if 안내['집계밖'] and not 안내.get('패키지형'):
            경고.append(f"「월별 실적집계」에 없는 활동세부가 붙은 줄이 {안내['집계밖']:,}건 있습니다. "
                        f"그대로 두면 그 금액이 집계표에서 빠집니다.")
        if 안내['매핑밖']:
            경고.append(f"BS_IS_매핑 표에 없는 계정이 {안내['매핑밖']:,}건 있습니다. "
                        f"계정과목·계정분류가 빈칸으로 나옵니다.")
        if 안내.get('수식고침'):
            경고.append(f"「{안내['당월시트']}」 {' · '.join(안내['수식고침'])} 칸이 "
                        f"빈칸(G5·H5)을 보고 있어 매출원가 누적이 잘못 나오던 것을 "
                        f"기준 월($H$2)을 보도록 고쳤습니다.")
        for 말 in 경고:
            st.caption(f'※ {말}')

    # ── 자료실에 쌓인 연결패키지 목록
    if 패키지들:
        줄 = ''.join(
            f'<tr class="{"total" if i == 0 else "sub"}">'
            f'<td class="name">{d["월"]}월말</td>'
            f'<td class="name">v{d["버전"]}{"  ← 최신" if i == 0 else ""}</td>'
            f'<td class="lft">{d["이름"]}</td>'
            f'<td>{datetime.datetime.fromtimestamp(d["시각"]):%Y-%m-%d %H:%M}</td>'
            f'<td>{d["크기"]/1e6:,.1f}MB</td></tr>'
            for i, d in enumerate(패키지들[:12]))
        st.html(f"""<div class="wrap" translate="no">
  <div class="card" style="margin-top:14px">
    <h3>자료실에 보관된 연결패키지</h3>
    <div class="sub">{os.path.abspath(자료실)} · 새 것부터 · 번호는 같은 실적 월 안에서 v1 부터 셉니다</div>
    <table class="lined" style="margin-top:10px">
      <colgroup><col style="width:10%"><col style="width:10%">
                <col style="width:49%"><col style="width:19%"><col style="width:12%"></colgroup>
      <thead><tr><th>실적 월</th><th>버전</th><th>파일 이름</th><th>만든 때</th><th>크기</th></tr></thead>
      <tbody>{줄}</tbody>
    </table>
  </div>
</div>""")

    # ── 자료실에 쌓인 버전 목록
    if 보관목록:
        줄 = ''.join(
            f'<tr class="{"total" if i == 0 else "sub"}">'
            f'<td class="name">{_달표시(d)}</td>'
            f'<td class="name">{_만든날표시(d)}</td>'
            f'<td class="name">v{d["버전"]}{"  ← 최신" if i == 0 else ""}</td>'
            f'<td class="lft">{d["이름"]}</td>'
            f'<td>{datetime.datetime.fromtimestamp(d["시각"]):%Y-%m-%d %H:%M}</td>'
            f'<td>{d["크기"]/1e6:,.1f}MB</td></tr>'
            for i, d in enumerate(보관목록[:12]))
        st.html(f"""<div class="wrap" translate="no">
  <div class="card" style="margin-top:14px">
    <h3>자료실에 보관된 실적보고 엑셀</h3>
    <div class="sub">{os.path.abspath(자료실)} · 새 것부터 ·
      번호는 만든 날마다 v1 부터 셉니다 (웹사이트를 업데이트하면 이 목록은 비워집니다)</div>
    <table class="lined" style="margin-top:10px">
      <colgroup><col style="width:8%"><col style="width:12%"><col style="width:8%">
                <col style="width:41%"><col style="width:19%"><col style="width:12%"></colgroup>
      <thead><tr><th>실적 월</th><th>만든 날</th><th>버전</th><th>파일 이름</th>
                 <th>만든 때</th><th>크기</th></tr></thead>
      <tbody>{줄}</tbody>
    </table>
  </div>
</div>""")
    st.stop()


# ★화면은 언제나 「마지막으로 올린 실적 엑셀」을 기준으로 그립니다★
if 업로드파일 is not None:
    데이터바이트 = 업로드파일.getvalue()
    with open(CACHE_PATH, 'wb') as f:
        f.write(데이터바이트)
    try:
        with open(이름기록, 'w', encoding='utf-8') as f:
            f.write(업로드파일.name)
    except OSError:
        pass
    쓰는파일 = 업로드파일.name
    설정칸.caption(f'✔ 방금 올림: {업로드파일.name}')
elif os.path.exists(CACHE_PATH):
    with open(CACHE_PATH, 'rb') as f:
        데이터바이트 = f.read()
    쓰는파일 = 쓰던파일이름()
    설정칸.caption(f'📁 사용 중: {쓰는파일}' if 쓰는파일 else '📁 마지막으로 올린 자료를 쓰고 있습니다')
else:
    # 자료가 없을 때 — 여기서 바로 올릴 수 있게 합니다 (사이드바까지 가지 않도록)
    st.html(f"""<div class="wrap" translate="no">
  <div class="card" style="max-width:820px; margin-top:8px">
    <div class="page-head" style="margin-bottom:6px">
      <span class="t">자료를 올려 주세요</span></div>
    <div style="font-size:15px; line-height:1.85; color:{T['ink2']}; margin-top:10px">
      아래 상자에 파일을 <b style="color:{T['ink']}">끌어다 놓기만</b> 하시면 됩니다.
      여러 개를 한꺼번에 올리셔도 파일 안을 보고 알아서 제자리에 넣습니다.
      <div style="margin-top:10px; font-size:14px; line-height:1.9">
        · <b style="color:{T['ink']}">실적보고 엑셀</b> (「26년 원장_raw」 시트가 든 파일)
          — 이 화면의 표와 그래프를 그립니다<br>
        · <b style="color:{T['ink']}">퀵북 원장 원본</b> (General Ledger)
          — 실적보고 엑셀작성 · 연결패키지에 씁니다<br>
        · <b style="color:{T['ink']}">연결패키지 양식</b> (본사에서 받은 원본)
          — 연결패키지 재무제표를 만듭니다
      </div>
    </div>
  </div>
</div>""")
    첫칸, _빈 = st.columns([2, 1], gap='small')
    with 첫칸:
        올린것 = st.file_uploader('자료 올리기 (.xlsx · 여러 개 가능)', type=['xlsx'],
                                  accept_multiple_files=True, key='첫화면업로드')
    if 올린것:
        실적, 목록 = 자료받기(올린것)
        for _n, _k in 목록:
            st.caption(f'{"✔" if _k else "?"} {_n} → '
                       + 종류이름.get(_k, '무슨 자료인지 모르겠습니다'))
        if 실적 is not None:
            try:
                with open(CACHE_PATH, 'wb') as f:
                    f.write(실적.getvalue())
                with open(이름기록, 'w', encoding='utf-8') as f:
                    f.write(실적.name)
                보관해두기(실적.getvalue(), 실적.name)
            except OSError as e:
                st.error(f'저장하지 못했습니다: {e}')
            else:
                st.rerun()
        else:
            st.caption('※ 표를 그리려면 「26년 원장_raw」 시트가 든 '
                       '실적보고 엑셀이 하나는 있어야 합니다.')
    st.html(f"""<div class="wrap" translate="no">
  <div class="card" style="max-width:820px; margin-top:4px">
    <div style="font-size:13.5px; line-height:1.85; color:{T['ink3']}">
      <b style="color:{T['ink2']}">왜 다시 올려야 하나요?</b><br>
      회사 재무자료는 인터넷(GitHub)에 올리지 않기 때문에, 웹사이트에는 올려 주신 파일이
      <b style="color:{T['ink2']}">잠시 머물다 지워집니다</b> — 웹사이트를 업데이트했을 때,
      그리고 며칠 아무도 안 들어와 서버가 잠들었다 깨어날 때 지워집니다.
      평소 쓰실 때는 지워지지 않습니다.<br>
      늘 켜 두고 쓰시려면 <b style="color:{T['ink2']}">「1. 대시보드_실행.bat」</b> 로
      회사 PC에서 여시면 됩니다 — 그때는 파일이 D:\\5. CTKonnect\\11-1. CTKonnect-Dashboard 에 그대로 남아
      다시 올리실 필요가 없습니다.
    </div>
  </div>
</div>""")
    st.stop()

법인 = 'CTKonnect'
if 메뉴 == 'CTKonnect':              # 2026-11-03 : 퀵북 원장 · 연결패키지를 함께 올리는 칸
    with st.expander('CTKonnect 자료 올리기 — 퀵북 원장 · 연결패키지 (함께 올려도 됩니다)', expanded=False):
        _uss = st.file_uploader('퀵북 원장(General Ledger) · CTKonnect 연결패키지 (.xlsx · 여러 개 가능)', type=['xlsx'],
                                accept_multiple_files=True, key='us_files')
        st.caption('원장 + 연결패키지 → 그 패키지 양식으로 원장을 반영 · 연결패키지만 → 그대로 · 원장만 → 지금 쓰는 패키지 양식으로 반영 '
                   '(환율 · 통장 대사가 필요하면 「실적보고 엑셀작성」 을 쓰세요)')
        if _uss:
            US올린것처리(_uss)
    _안 = st.session_state.get('_US올림안내')
    if _안:
        (st.success if _안[0] else st.warning)(_안[1])
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, 'rb') as _f:
            데이터바이트 = _f.read()
        쓰는파일 = 쓰던파일이름() or 쓰는파일
if 메뉴 == '씨티케이커넥트':         # 2026-11-03 : 한국 법인 — 같은 화면을 씨티케이커넥트 연결패키지(KRW)로 그립니다
    법인 = '씨티케이커넥트'
    with st.expander('씨티케이커넥트 자료 올리기 — 더존 분개장 · 연결패키지 (함께 올려도 됩니다)', expanded=not _KR패키지바이트(분개장보기=False)[0]):
        _krs = st.file_uploader('더존 분개장 · 씨티케이커넥트 연결패키지 (.xlsx · 여러 개 가능)', type=['xlsx'],
                                accept_multiple_files=True, key='kr_files')
        st.caption('분개장 + 연결패키지 → 그 패키지 양식으로 분개장을 반영 · 연결패키지만 → 그대로 · 분개장만 → 지금 쓰는 패키지 양식으로 반영')
        _krs_도장 = tuple(sorted((f.name, int(getattr(f, 'size', 0) or 0)) for f in (_krs or [])))
        if _krs and st.session_state.get('_kr_files_도장') != _krs_도장:
            KR올린것처리(_krs)
            st.session_state['_kr_files_도장'] = _krs_도장
    _kb, _kn = _KR패키지바이트()
    _안 = st.session_state.get('_KR올림안내')
    if _안:
        (st.success if _안[0] else st.warning)(_안[1])
    if not _kb:
        st.html(f"""<div class="wrap" translate="no">
  <div class="page-head"><span class="t">실적보고 — 씨티케이커넥트</span><span class="pill">자료 없음</span></div>
  <div class="card" style="max-width:820px; margin-top:8px">
    <div style="font-size:15px; line-height:1.85; color:{T['ink2']}">
      씨티케이커넥트(한국 법인) 연결패키지가 아직 없습니다.<br>
      위 「씨티케이커넥트 자료 올리기」 에 <b style="color:{T['ink']}">더존 분개장</b>과
      <b style="color:{T['ink']}">「…연결패키지_씨티케이커넥트_○월말_v○.xlsx」</b> 를 함께 올려 주세요.
      한 번 올린 뒤에는 분개장만 올려도 그 패키지 양식으로 바꿔 보여 드립니다.
      PC 에서는 자료실 달 폴더의 「씨티케이커넥트」 하위 폴더에 넣어 두면 자동으로 읽습니다.</div></div></div>""")
        st.stop()
    데이터바이트, 쓰는파일 = _kb, _kn
    설정칸.caption(f'📁 씨티케이커넥트 자료: {_kn}')

# ══════════════════════════════════════════════════════════════
# 실적보고 → 연결 손익계산서 (2026-11-02)
#   ① CTKonnect(미국 · USD) → ② 씨티케이커넥트(한국 · KRW) → ③ 연결 (단순합산 · 내부거래 제거 · 연결)
#   두 법인 모두 연결패키지(26년 원장_raw · BS_IS_매핑) 양식에서 읽습니다.
# ══════════════════════════════════════════════════════════════
if 메뉴 == '연결 손익계산서':
    미국이름 = ['CTKONNECT', 'CTK ONNECT']           # 한국 원장에서 미국 법인을 찾을 글자 (더존 거래처명 「CTK onnect」)
    한국이름 = ['컴플리톤', '씨티케이커넥트', 'CTKCONNECT KOREA']             # 미국 원장에서 한국 법인을 찾을 글자 (구 이름 컴플리톤 포함)

    def _kr자료실():
        후보 = [f for f in glob.glob(os.path.join(자료실, '**', '*.xlsx'), recursive=True)
                if '씨티케이커넥트' in os.path.basename(f) and '연결패키지' in os.path.basename(f)
                and not os.path.basename(f).startswith('~$')]
        return max(후보, key=os.path.getmtime) if 후보 else None

    st.html(f"""<div class="wrap" translate="no">
  <div class="page-head"><span class="t">실적보고 — 연결 손익계산서</span><span class="pill">CTKonnect + 씨티케이커넥트</span></div>
  <div class="sub" style="margin:2px 0 6px">연결정산표 — CTKonnect · 씨티케이커넥트 · 합계 · 연결조정 · 연결재무 (누적)</div></div>""")

    with st.expander('씨티케이커넥트 연결패키지 올리기', expanded=False):
        _kr = st.file_uploader('씨티케이커넥트(한국 법인) 연결패키지 — 「26년 원장_raw」 · 「BS_IS_매핑」 시트가 든 엑셀',
                               type=['xlsx'], key='kr_pkg')
        if _kr is not None:
            try:
                with open(KR보관, 'wb') as _f:
                    _f.write(_kr.getvalue())
                with open(KR이름표, 'w', encoding='utf-8') as _f:
                    _f.write(_kr.name)
            except OSError:
                pass
        st.caption('PC 에서는 자료실에 「…연결패키지_씨티케이커넥트_○월말_v○.xlsx」 를 넣어 두면 자동으로 읽습니다.')

    KR바이트, KR파일 = _KR패키지바이트()

    def _읽기(바이트):
        r = 원장읽기(바이트)
        return r[0], r[2]
    try:
        US원장, US연도 = _읽기(데이터바이트)
    except Exception as e:
        st.error(f'CTKonnect 연결패키지를 읽지 못했습니다 — {e}')
        st.stop()
    KR원장, KR오류 = None, ''
    if KR바이트:
        try:
            KR원장, _ = _읽기(KR바이트)
        except Exception as e:
            KR오류 = str(e)

    # 평균환율 (원/USD) — CTKonnect 연결패키지 IS M5
    try:
        _w = openpyxl.load_workbook(io.BytesIO(데이터바이트), read_only=True, data_only=False)
        _기본환율 = float(_w['IS']['M5'].value)
    except Exception:
        _기본환율 = 0.0
    c1, c2, c3 = st.columns([1.2, 1, 2.4], gap='small')
    통화선택 = c1.radio('보여 줄 통화', ['원화 (KRW)', '달러 (USD)'], horizontal=True, key='conso_ccy')
    환율 = c2.number_input('평균환율 (원/USD)', min_value=0.0, value=round(_기본환율, 2), step=0.01, format='%.2f', key='conso_fx',
                          help='기본값 = CTKonnect 연결패키지 IS 시트 M5 (1~결산월 평균환율)')
    원화 = 통화선택.startswith('원화')
    단위 = 'KRW' if 원화 else 'USD'
    if 환율 <= 0:
        st.warning('평균환율이 비어 있습니다 — 환율을 넣어야 두 법인을 같은 통화로 더할 수 있습니다.')

    def _월끝(원장):
        x = 원장[원장['계정분류'].eq('IS') & 원장['금액'].abs().gt(0.005)]
        return int(x['월'].max()) if len(x) else 0
    US끝 = _월끝(US원장)
    KR끝 = _월끝(KR원장) if KR원장 is not None else 0
    끝월 = min(US끝, KR끝) if KR원장 is not None else US끝
    US배수 = (환율 if 원화 else 1.0) if (환율 > 0 or not 원화) else 0.0
    KR배수 = 1.0 if 원화 else (1.0 / 환율 if 환율 > 0 else 0.0)

    # ── 손익계산서 과목표 : CTKonnect 연결패키지 IS 시트 (Ⅰ.~ 구역 · 1. 2. 과목) 그대로
    _w = openpyxl.load_workbook(io.BytesIO(데이터바이트), read_only=True)
    _is = _w['IS']
    구역 = []                     # [(구역이름, 성격, [과목...])]
    for r in _is.iter_rows(min_row=8, max_row=120, max_col=5, values_only=True):
        로마, 번호, 이름 = r[1], r[3], r[4]
        이름 = str(이름 or '').replace(' ', '').strip()
        if 로마 and 이름:
            구역.append([이름, None, []])
        elif 번호 and 이름 and 구역:
            구역[-1][2].append(이름)
    for z in 구역:                # 과목이 따로 없는 구역(법인세비용 등)은 구역 이름 자체가 과목
        if not z[2]:
            z[2] = [z[0]]
    계산줄 = {'매출총이익', '영업이익', '법인세비용차감전계속사업이익', '계속사업이익', '당기순이익(손실)', '중단사업이익', '기타포괄손익'}
    수익구역 = {'매출액', '기타수익', '금융수익', '지분법손익'}

    def _과목금액(원장, 배수):
        """과목별 1~끝월 누적 — 수익은 +, 비용은 + (표시용) · 외환 · 환산손익은 순액으로 이익/손실 칸에 나눔."""
        if 원장 is None:
            return {}
        d = 원장[원장['계정분류'].eq('IS') & 원장['월'].between(1, 끝월)]
        합 = d.groupby(d['계정과목'].astype(str).str.replace(' ', ''))['금액'].sum() * 배수     # 대변 − 차변
        o = {}
        for k, v in 합.items():
            o[k] = o.get(k, 0.0) + float(v)
        for 순, 이익, 손실 in (('외환차손익', '외환차익', '외환차손'), ('외화환산손익', '외화환산이익', '외화환산손실')):
            n = o.pop(순, 0.0) + o.pop(이익, 0.0) + o.pop(손실, 0.0)
            o[이익], o[손실] = max(n, 0.0), max(-n, 0.0) * -1     # 손실은 대변−차변이 음수
        return o
    US금 = _과목금액(US원장, US배수)
    KR금 = _과목금액(KR원장, KR배수)

    # ── 내부거래 (연결조정 하나) : 손익 계정 줄 중 거래처가 상대 법인인 줄
    def _내부(원장, 상대이름, 법인, 배수):
        if 원장 is None:
            return pd.DataFrame()
        d = 원장[원장['계정분류'].eq('IS') & 원장['월'].between(1, 끝월)].copy()
        n = d['Name'].fillna('').astype(str).str.upper()
        d = d[n.apply(lambda s: any(k.upper() in s for k in 상대이름))].copy()
        d['법인'] = 법인
        d['환산'] = d['금액'] * 배수
        d['과목'] = d['계정과목'].astype(str).str.replace(' ', '')
        return d
    IC = pd.concat([_내부(US원장, 한국이름, 'CTKonnect', US배수), _내부(KR원장, 미국이름, '씨티케이커넥트', KR배수)])
    차변, 대변, 비고 = {}, {}, {}
    가정글 = ''
    if KR원장 is not None and len(IC):
        과목구역 = {과: 구 for 구, _, 목 in 구역 for 과 in 목}
        원가짝 = {'상품매출': '상품매출원가', '제품매출': '제품매출원가', '용역매출': '용역매출원가', '기타매출': '기타매출원가'}
        가정들, 확인들 = [], []
        for 판매, 구매 in (('씨티케이커넥트', 'CTKonnect'), ('CTKonnect', '씨티케이커넥트')):
            s_ = IC[IC['법인'] == 판매]; b_ = IC[IC['법인'] == 구매]
            수익 = s_[s_['과목'].map(lambda k: 과목구역.get(k, '') in 수익구역)]
            비용 = b_[b_['과목'].map(lambda k: 과목구역.get(k, '') not in 수익구역)]
            수익합 = {k: float(v) for k, v in 수익.groupby('과목')['환산'].sum().items() if abs(v) >= 0.5}
            비용합 = {k: -float(v) for k, v in 비용.groupby('과목')['환산'].sum().items() if abs(v) >= 0.5}
            for k, v in 수익합.items():                    # 파는 쪽 내부 수익 지움 = 차변
                차변[k] = 차변.get(k, 0.0) + v
            한도 = sum(수익합.values())
            for k, v in 비용합.items():                    # 사는 쪽 내부 비용 지움 = 대변 (파는 쪽 내부 수익만큼만)
                쓸 = min(v, max(한도, 0.0)); 한도 -= 쓸
                if 쓸 >= 0.5:
                    대변[k] = 대변.get(k, 0.0) + 쓸
                if v - 쓸 >= 0.5:
                    확인들.append(f'{구매} 의 {k} {v - 쓸:,.0f} 은 {판매} 쪽 내부 수익이 없어 지우지 않았습니다 — 확인 필요')
            if 한도 >= 0.5:                                  # 사는 쪽 원장에서 비용을 못 찾은 내부 매출 → 같은 금액을 파는 쪽 매출원가에서 지움
                for k, v in 수익합.items():
                    if 과목구역.get(k) != '매출액' or 한도 < 0.5:
                        continue
                    쓸 = min(v, 한도); 한도 -= 쓸
                    짝 = 원가짝.get(k, '상품매출원가')
                    대변[짝] = 대변.get(짝, 0.0) + 쓸
                    비고[짝] = f'「가정」 내부매출과 같은 금액을 매출원가에서 지움 (재고에 남은 내부이익은 그대로)'
                    가정들.append(f'{판매} → {구매} 내부매출 {k} {쓸:,.0f}')
        if 가정들:
            가정글 = ('<div class="sub" style="color:#C00000; font-weight:700; margin-top:8px">「가정」 '
                      + ' · '.join(가정들) + ' 은 사는 법인 원장에서 상대 이름으로 매출원가를 찾지 못해, 내부매출과 같은 금액을 '
                      '매출원가에서 지웠습니다. 사는 법인 재고(상품 · 미착품 · 선급금)에 아직 남은 물건의 내부이익은 빼지 않았습니다 — '
                      '재고 자료가 있어야 정확히 지울 수 있습니다.</div>')
        if 확인들:
            가정글 += '<div class="sub" style="color:#C00000; margin-top:4px">' + ' · '.join(확인들) + '</div>'
        for k in set(차변) | set(대변):
            비고.setdefault(k, '내부거래 제거')

    def _칸(v):
        if v is None or pd.isna(v) or abs(v) < 0.5:
            return '-'
        t = f'{abs(v):,.0f}'
        return f'<span class="neg">({t})</span>' if v < 0 else t

    # ── 표 줄 만들기 : 구역 · 과목 (금액 있는 과목만) · 계산줄
    줄들 = []        # (이름, 수준, us, kr, 차, 대, 비고)
    구역합 = {}
    def _과목줄(과, 수익):
        u = US금.get(과, 0.0); k = KR금.get(과, 0.0)
        if not 수익:
            u, k = -u, -k
        차, 대 = 차변.get(과, 0.0), 대변.get(과, 0.0)
        연 = (u + k - 차 + 대) if 수익 else (u + k + 차 - 대)
        return u, k, 차, 대, 연
    쓴과목 = set()
    for 이름, _, 목 in 구역:
        if 이름 in 계산줄:
            continue
        수익 = 이름 in 수익구역
        합 = [0.0] * 5
        자식 = []
        for 과 in 목:
            v = _과목줄(과, 수익)
            쓴과목.add(과)
            합 = [a + b for a, b in zip(합, v)]
            if any(abs(x) >= 0.5 for x in v) and 과 != 이름:
                자식.append((과, 1, *v, 비고.get(과, '')))
        구역합[이름] = 합
        if any(abs(x) >= 0.5 for x in 합) or 이름 in ('매출액', '매출원가', '판매비와관리비'):
            줄들.append((이름, 0, *합, 비고.get(이름, '') if 목 == [이름] else ''))
            줄들.extend(자식)
        # 계산줄은 해당 구역 뒤에 끼움
        def _계(n, 값):
            줄들.append((n, 0, *값, ''))
        if 이름 == '매출원가':
            _계('매출총이익', [구역합['매출액'][0] - 합[0], 구역합['매출액'][1] - 합[1], 0, 0, 구역합['매출액'][4] - 합[4]])
        if 이름 == '판매비와관리비':
            gp = [구역합['매출액'][i] - 구역합['매출원가'][i] for i in (0, 1, 4)]
            _계('영업이익', [gp[0] - 합[0], gp[1] - 합[1], 0, 0, gp[2] - 합[4]])
    def _g(n, i):
        return 구역합.get(n, [0.0] * 5)[i]
    세전 = [(_g('매출액', i) - _g('매출원가', i) - _g('판매비와관리비', i) + _g('기타수익', i) - _g('기타비용', i)
             + _g('금융수익', i) - _g('금융비용', i) + _g('지분법손익', i)) for i in range(5)]
    i법 = next((j for j, z in enumerate(줄들) if z[0] == '법인세비용'), None)
    세전줄 = ('법인세비용차감전이익', 0, 세전[0], 세전[1], 0, 0, 세전[4], '')
    if i법 is None:
        줄들.append(세전줄)
    else:
        줄들.insert(i법, 세전줄)
    순 = [세전[i] - _g('법인세비용', i) for i in range(5)]
    줄들.append(('당기순이익', 0, 순[0], 순[1], 0, 0, 순[4], ''))
    빠진 = sorted({k for k in list(US금) + list(KR금) if k not in 쓴과목 and abs(US금.get(k, 0) + KR금.get(k, 0)) >= 0.5})

    굵게 = {'매출총이익', '영업이익', '법인세비용차감전이익', '당기순이익'}
    몸 = ''
    for 이름, 수준, u, k, 차, 대, 연, 비 in 줄들:
        st_ = ' style="font-weight:700; background:rgba(127,127,127,.07)"' if 이름 in 굵게 else (' style="font-weight:700"' if 수준 == 0 else '')
        몸 += (f'<tr{st_}><td class="lft">{"&nbsp;&nbsp;&nbsp;" * 수준}{_html.escape(이름)}</td><td>{_칸(u)}</td>'
               f'<td>{_칸(k) if KR원장 is not None else "-"}</td><td>{_칸(u + k)}</td>'
               f'<td>{_칸(차)}</td><td>{_칸(대)}</td><td><b>{_칸(연)}</b></td>'
               f'<td class="lft" style="font-weight:400; color:{"#C00000" if "가정" in 비 else T["ink3"]}">{_html.escape(비)}</td></tr>')
    st.html(f"""<div class="wrap" translate="no"><div class="card" style="margin-top:10px">
      <h3>연결 손익계산서 — {US연도}년 1~{끝월}월 누적 <span class="unitbadge">단위 {단위}</span></h3>
      <div class="sub">평균환율 {환율:,.2f} 원/USD · {'CTKonnect(USD)를 원화로 환산' if 원화 else '씨티케이커넥트(KRW)를 달러로 환산'} ·
        합계 = 두 법인 단순합산 · 연결조정 = 두 법인 사이 거래(내부거래) 제거 · 연결재무 = 합계 ∓ 연결조정</div>
      <div class="scrollx"><table class="lined" style="width:100%; min-width:1080px">
        <colgroup><col style="width:17%"><col style="width:11%"><col style="width:11%"><col style="width:11%"><col style="width:10%"><col style="width:10%"><col style="width:11%"><col style="width:19%"></colgroup>
        <thead><tr><th class="lft" rowspan="2">과목</th><th rowspan="2">CTKonnect</th><th rowspan="2">씨티케이커넥트</th><th rowspan="2">합계</th>
          <th colspan="2">연결조정</th><th rowspan="2">연결재무</th><th rowspan="2">비고</th></tr>
          <tr><th>차변</th><th>대변</th></tr></thead>
        <tbody>{몸}</tbody></table></div>{가정글}
      {'' if not 빠진 else f'<div class="sub" style="color:#C00000; margin-top:6px">IS 과목표에 없는 과목이 있어 표에 넣지 못했습니다 : {" · ".join(빠진)} — BS_IS_매핑 확인 필요</div>'}
      {'' if KR원장 is not None else ('<div class="sub" style="margin-top:8px">' + (('씨티케이커넥트 연결패키지를 읽지 못했습니다 — ' + _html.escape(KR오류)) if KR오류 else '씨티케이커넥트 연결패키지가 아직 없어 CTKonnect 만 들어 있습니다 — 위 「씨티케이커넥트 연결패키지 올리기」 에서 올려 주세요.') + ' 내부거래는 두 법인 자료가 다 있을 때 지웁니다.</div>')}
      <div class="sub" style="margin-top:6px">CTKonnect : {_html.escape(str(쓰는파일 or ''))}{(' · 씨티케이커넥트 : ' + _html.escape(KR파일)) if KR파일 else ''}</div>
    </div></div>""")
    if len(IC):
        줄 = ''.join(
            f'<tr><td class="lft">{r["법인"]}</td><td>{r["거래일"]:%Y-%m-%d}</td><td class="lft">{_html.escape(str(r["계정과목"]))}</td>'
            f'<td class="lft">{_html.escape(str(r["Name"]))}</td><td class="lft" style="font-weight:400">{_html.escape(str(r["Description"] or ""))[:60]}</td>'
            f'<td>{_칸(r["환산"])}</td></tr>' for _, r in IC.sort_values(['법인', '거래일']).iterrows())
        with st.expander(f'내부거래 내역 — 두 법인 사이 손익 거래 {len(IC)}줄', expanded=False):
            st.html(f"""<div class="wrap" translate="no"><div class="card">
              <div class="sub">손익 계정 줄 중 거래처가 상대 법인(한국 : {' · '.join(한국이름)} / 미국 : CTKonnect, Inc.)인 줄 · 금액 = 이익(+) · 비용(−) · 단위 {단위}</div>
              <div class="scrollx"><table class="lined" style="width:100%; min-width:820px">
                <thead><tr><th class="lft">법인</th><th>일자</th><th class="lft">계정</th><th class="lft">거래처</th><th class="lft">적요</th><th>금액</th></tr></thead>
                <tbody>{줄}</tbody></table></div></div></div>""")
    st.stop()

try:
    (당해원장, 전년원장, 당해연도, 매핑표, 기초현금,
     입력값, BS밑자료) = 원장읽기(데이터바이트)
except Exception as e:
    st.error(f'파일을 읽지 못했습니다: {e}')
    st.stop()

전년연도 = 당해연도 - 1
전년없음 = len(전년원장) == 0      # 2026-11-02 : 전년 원장이 없으면 전년 칸 · 연도 고르기를 숨김 (26년은 25년 없이, 27년부터 전년 비교)
P = 손익표(당해원장, '활동')          # 대시보드·월간실적은 활동분류 기준을 씁니다
P전 = 손익표(전년원장, '활동')
P손익 = 손익표(당해원장, '손익')      # 실적집계 화면에서 고를 수 있는 손익계산서 기준
P손익전 = 손익표(전년원장, '손익')
보고월 = int(당해원장['월'].max())
# 사이드바 제목 아래 「최종자료 · 마지막 올린 사람」 — 세 법인 대시보드가 같은 모양입니다
_쪽지 = 자료기록읽기()
_올린이줄 = ''
if _쪽지.get('때'):
    _누가 = _쪽지.get('올린이') or '미선택'
    _올린이줄 = (f'<div style="margin-top:6px; padding-top:6px;'
                 f' border-top:1px solid rgba(255,255,255,.14);'
                 f' font-size:11px; color:rgba(255,255,255,.70)">'
                 f'마지막 올린 곳 <b style="color:#fff">{_누가}</b><br>{_쪽지["때"]}</div>')

자료월칸.markdown(
    '<div style="margin:-6px 0 10px; padding:9px 12px; border-radius:10px;'
    ' background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.18)">'
    '<div style="font-size:11px; letter-spacing:.06em;'
    ' color:rgba(255,255,255,.65)">최종자료</div>'
    f'<div style="font-size:16px; font-weight:700; color:#fff; line-height:1.35">'
    f'{당해연도}년 {보고월}월</div>'
    f'<div style="font-size:11px; color:rgba(255,255,255,.60)">'
    f'{당해연도}년 1~{보고월}월 누적</div>'
    f'{_올린이줄}</div>', unsafe_allow_html=True)
기준일 = 당해원장['거래일'].max()

_쓰는파일줄 = (f'<div title="{쓰는파일}">자료 <b>{쓰는파일[:26]}'
               f'{"…" if len(쓰는파일) > 26 else ""}</b></div>') if 쓰는파일 else ''
st.sidebar.markdown(
    f"""<div class="side-foot">
      <div>기준 <b>{당해연도}년 1~{보고월}월</b></div>
      <div>원장 {len(당해원장):,}건</div>
      {_쓰는파일줄}
    </div>""", unsafe_allow_html=True)

미분류건수 = int((당해원장['계정분류'].eq('IS') & 당해원장['분류'].eq('판관비')
                & 당해원장['활동세부'].isna()).sum())


def 누적(t, 연도표=None, upto=None):
    upto = upto or 보고월
    return t[list(range(1, upto + 1))].sum(axis=1)


당해누적 = 누적(P)
전년동기 = 누적(P전)
전년연간 = P전[list(range(1, 13))].sum(axis=1)




# 오른쪽 세부내역 패널에서 항목 하나를 고르면 어떤 자금분류(구분) 코드를 찾아볼지 매핑.
# 부호구분 'DEBT'는 같은 코드를 부호로 다시 나눕니다 — '+' 는 차입(들어온 돈), '−' 는 상환(나간 돈).
현금흐름목록 = [
    ('매출채권 회수', 'AR', None),
    ('기타입금', 'ETCIN', None),
    ('매입채무', 'AP', None),
    ('급여', 'PAY', None),
    ('기타 비용', 'OTHOP', None),
    ('설비 투자', 'CAPEX', None),
    ('금융자산 투자', 'INV', None),
    ('유상증자', 'EQ', None),
    ('차입', 'DEBT', '+'),
    ('차입 상환', 'DEBT', '-'),
    ('리스료 지급', 'LEASE', None),
    ('환율변동 효과', 'FX', None),
]

SRED = SRED월별(당해원장)
현금잔고, 현금증감 = 현금잔고월별(당해원장, 기초현금)
자금표, 자금상세 = 자금변동(당해원장)


# CTKonnect 장부 통화는 USD 하나 — 표시 통화 · 환율 칸이 없습니다 (캐나다는 CAD/USD 선택)
통화 = 'KRW' if 법인 == '씨티케이커넥트' else 'USD'      # 2026-11-03 : 씨티케이커넥트는 원화
환율 = 1.0
if 메뉴 in ('월간실적', 'CTKonnect', '씨티케이커넥트') and not 입력값['단일'] and not 입력값['서술']:
    설정칸.caption('ℹ️ 연결패키지에 「월간보고_입력」 시트가 없어 목표·진척률·코멘트는 비어 있습니다.')
배율 = 1.0
금액나눔 = 표시나눔 = 1000 if 법인 == '씨티케이커넥트' else 1      # CTKonnect : USD 그대로 (2026-10-29) · 씨티케이커넥트 : 천원 (2026-11-03)


def 상품채권html(원장, 기준일):
    """상품매출 청구 건별 미수 · 경과 구간 (대출금 내역 표 모양)."""
    상품전표 = set(원장.loc[원장['계정과목'].astype(str).eq('상품매출'), 'TransactionID'])
    ar = 원장[원장['계정과목'].astype(str).eq('매출채권')].copy()
    if not 상품전표 or ar.empty:
        return ''
    이름 = lambda n: re.sub(r'\s*\((IC\s+)?(Customer|Customers|Vendor)\)\s*$', '',
                            '' if n is None or (isinstance(n, float) and pd.isna(n)) else str(n)).strip() or '(거래처 없음)'
    ar['거래처'] = ar['Name'].map(이름)
    대상 = set(ar.loc[ar['TransactionID'].isin(상품전표) & ar['Debit'].gt(0), '거래처'])
    상품매출액 = (원장[원장['계정과목'].astype(str).eq('상품매출')]
                  .assign(v=lambda t: t['Credit'] - t['Debit']).groupby('TransactionID')['v'].sum())
    줄들, 합 = [], {'청구': 0.0, '회수': 0.0, '미수': 0.0}
    구간합 = {c: 0.0 for c in 채권구간}
    for 곳, x in ar[ar['거래처'].isin(대상)].groupby('거래처'):
        x = x.sort_values(['거래일', 'TransactionID'], kind='stable')
        큐 = []
        for _, r in x.iterrows():
            if r['Debit'] > 0:
                큐.append({'일자': r['거래일'], '청구': float(r['Debit']), '회수': 0.0, 'TID': r['TransactionID'],
                           '적요': str(r['Description'] or ''), '상품': r['TransactionID'] in 상품전표})
            돈 = float(r['Credit'])
            for q in 큐:
                if 돈 <= 1e-9:
                    break
                뺄 = min(돈, q['청구'] - q['회수'])
                q['회수'] += 뺄
                돈 -= 뺄
        for q in 큐:
            if not q['상품']:
                continue
            미수 = q['청구'] - q['회수']
            만기 = _만기일(q['일자'])
            경과 = (기준일 - 만기).days
            구 = _채권구간(경과) if 미수 > 0.005 else ''
            if 구:
                구간합[구] += 미수
            for k, v in (('청구', q['청구']), ('회수', q['회수']), ('미수', 미수)):
                합[k] += v
            줄들.append((q['일자'], 곳, float(상품매출액.get(q['TID'], 0.0)), q['청구'], q['회수'], 미수, 만기, 경과, 구, q['적요']))
    if not 줄들:
        return ''
    줄들.sort(key=lambda t: (t[0], t[1]))
    몸 = ''.join(
        f'<tr class="sub"><td>{d:%Y-%m-%d}</td><td class="lft">{_html.escape(곳)}</td><td>{금액(매)}</td>'
        f'<td>{금액(청)}</td><td>{금액(회)}</td><td><b>{금액(미)}</b></td><td>{만:%m-%d}</td>'
        f'<td>{(경 if 미 > 0.005 and 경 > 0 else "-")}</td><td style="text-align:center">{구 or "회수 완료"}</td>'
        f'<td class="lft" style="font-weight:400;background:transparent">{_html.escape(적)}</td></tr>'
        for d, 곳, 매, 청, 회, 미, 만, 경, 구, 적 in 줄들)
    몸 += (f'<tr class="total"><td class="lft" colspan="2">합계</td><td>{금액(sum(t[2] for t in 줄들))}</td><td>{금액(합["청구"])}</td><td>{금액(합["회수"])}</td>'
           f'<td><b>{금액(합["미수"])}</b></td><td></td><td></td><td></td><td></td></tr>')
    구간글 = ' · '.join(f'{c} <b>{금액(v)}</b>' for c, v in 구간합.items() if abs(v) > 0.5) or '미수 없음'
    return f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card" style="margin-top:14px"><h3>상품매출 미수채권 Aging — 건별 <span class="unitbadge">단위 USD</span></h3>
        <div class="sub">상품매출(51101)이 잡힌 전표의 매출채권 청구만 · 청구액 = 마켓 수수료를 뺀 정산액 · 회수는 거래처별 오래된 청구부터 ·
          만기 = 청구한 달의 다음 달 말일 · 경과일 = 만기에서 {기준일:%Y-%m-%d} 까지 · 남은 미수 : {구간글}</div>
        <div class="scrollx"><table class="lined" style="width:100%; table-layout:fixed;">
          <colgroup><col style="width:9%"><col style="width:15%"><col style="width:9%"><col style="width:9%"><col style="width:9%">
            <col style="width:9%"><col style="width:6%"><col style="width:6%"><col style="width:8%"><col style="width:20%"></colgroup>
          <thead><tr><th>청구일</th><th>거래처</th><th>상품매출</th><th>청구액</th><th>회수</th><th>미수</th><th>만기</th>
            <th>경과일</th><th>구간</th><th>적요</th></tr></thead>
          <tbody>{몸}</tbody></table></div>
      </div></div>"""


대출한도_확인 = {k: float(v) for k, v in dict(_비밀값('loan_limit', {})).items()}     # 대출 한도 — 담당자 확인 (Secrets)
대출기초_확인 = {k: float(v) for k, v in dict(_비밀값('loan_opening', {})).items()}   # 25년 말 차입 잔액 — 사용자 확인 (Secrets) · 대출금 표에만 씀


def 대출금html(원장, 보고월, 전년=None, 기초보정=None, 한도보정=None, 기초설명=None, 통화='USD'):
    """차입금 계정(단기 · 장기)의 대출처별 차입 · 상환 · 잔액 · 이자와 거래 내역 (CTKonnect, 2026-10-29)."""
    d = 원장[원장['계정과목'].astype(str).str.contains('차입금') & 원장['월'].le(보고월)].copy()
    if d.empty:
        return ''
    이름 = lambda n: re.sub(r'\s*\((IC\s+)?(Customer|Customers|Vendor)\)\s*$', '',
                            '' if n is None or (isinstance(n, float) and pd.isna(n)) else str(n)).strip() or '(대출처 없음)'
    d['대출처'] = d['Name'].map(이름)
    # 기초 = 25년 원장_raw 의 차입금 (대변 − 차변) — 대출처별
    기초별 = pd.Series(dtype=float)
    if 전년 is not None and len(전년):
        p = 전년[전년['계정과목'].astype(str).str.contains('차입금')].copy()
        if len(p):
            기초별 = (p['Credit'] - p['Debit']).groupby(p['Name'].map(이름)).sum()
    확인기초 = False
    _기초확인 = 대출기초_확인 if 기초보정 is None else 기초보정
    _한도확인 = 대출한도_확인 if 한도보정 is None else 한도보정
    if 기초별.abs().sum() < 0.5 and _기초확인:     # 25년 원장이 없으면 담당자가 알려 준 25년 말 잔액
        기초별 = pd.Series(_기초확인, dtype=float)
        확인기초 = True
    이자 = 원장[원장['계정과목'].astype(str).eq('이자비용') & 원장['월'].le(보고월)].copy()
    이자['대출처'] = 이자['Name'].map(이름)
    이자별 = (이자['Debit'] - 이자['Credit']).groupby(이자['대출처']).sum()
    요약 = ''
    합 = [0.0, 0.0, 0.0, 0.0, 0.0]
    합한도 = [0.0, 0.0]
    곳들 = list(dict.fromkeys(list(기초별[기초별.abs() > 0.5].index) + list(d['대출처'])))
    for 곳 in 곳들:
        x = d[d['대출처'] == 곳]
        차입, 상환 = float(x['Credit'].sum()), float(x['Debit'].sum())
        기초 = float(기초별.get(곳, 0.0))
        잔 = 기초 + 차입 - 상환
        이 = float(이자별.get(곳, 0.0))
        율 = re.findall(r'(\d+\.\d+)\s*%', ' '.join(x['Description'].dropna().astype(str)))
        비고 = (f'상환이 기초 + 차입보다 {금액(-잔)} 많음 — 확인 필요' if 잔 < -0.5 else
                ((기초설명 or '기초 = 25년 말 차입 잔액 (담당자 확인 · 26년 원장에는 없음)') if 확인기초 and abs(기초) > 0.5 else
                 (f'최근 금리 {율[-1]}%' if 율 else '')))
        한도 = _한도확인.get(곳)
        남은 = (한도 - 잔) if 한도 else None
        for i, v in enumerate((기초, 차입, 상환, 잔, 이)):
            합[i] += v
        합한도[0] += 한도 or 0.0
        합한도[1] += 남은 or 0.0
        요약 += (f'<tr class="sub"><td class="lft">{_html.escape(곳)}</td><td>{금액(한도) if 한도 else "-"}</td>'
                 f'<td>{금액(기초)}</td><td>{금액(차입)}</td><td>{금액(상환)}</td>'
                 f'<td><b>{금액(잔)}</b></td><td>{금액(남은) if 남은 is not None else "-"}</td><td>{금액(이)}</td><td class="lft" style="font-weight:400;background:transparent">{_html.escape(비고)}</td></tr>')
    요약 += (f'<tr class="total"><td class="lft">합계</td><td>{금액(합한도[0])}</td>'
             + ''.join(f'<td>{금액(v)}</td>' for v in 합[:4]) + f'<td>{금액(합한도[1])}</td><td>{금액(합[4])}</td>'
             + '<td></td></tr>')
    내역 = ''
    for _, r in d.sort_values('거래일').iterrows():
        구분 = '차입' if r['Credit'] > 0 else '상환'
        내역 += (f'<tr class="sub"><td>{r["거래일"]:%Y-%m-%d}</td><td class="lft">{_html.escape(r["대출처"])}</td>'
                 f'<td>{구분}</td><td>{금액(r["Credit"] if 구분 == "차입" else -r["Debit"])}</td>'
                 f'<td class="lft" style="font-weight:400;background:transparent">{_html.escape(str(r["Description"] or ""))}</td></tr>')
    return f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>대출금 내역 <span class="unitbadge">단위 {단위K}</span></h3>
        <div class="sub">원장 차입금 계정 · 1~{보고월}월 · 잔액 = 기초 + 차입 − 상환 · 남은 한도 = 한도 − 잔액 · 한도는 담당자 확인 값 · 이자 = 원장 이자비용(대출처별)</div>
        <div class="scrollx"><table class="lined" style="width:100%; table-layout:fixed;">
          <colgroup>{'<col style="width:13%"><col style="width:6%"><col style="width:10.5%"><col style="width:10.5%"><col style="width:10.5%"><col style="width:10.5%"><col style="width:9%"><col style="width:12%"><col style="width:18%">' if 통화 == 'KRW' else '<col style="width:16%"><col style="width:9.5%"><col style="width:9%"><col style="width:9.5%"><col style="width:9.5%"><col style="width:9.5%"><col style="width:9.5%"><col style="width:8%"><col style="width:19.5%">'}</colgroup>
          <thead><tr><th>대출처</th><th>한도</th><th>기초 (25년 말)</th><th>차입</th><th>상환</th><th>{보고월}월 말 잔액</th>
            <th>남은 한도</th><th>이자 (1~{보고월}월)</th><th>비고</th></tr></thead>
          <tbody>{요약}</tbody></table></div>
        <div class="scrollx" style="margin-top:12px"><table class="lined" style="width:100%; table-layout:fixed;">
          <colgroup><col style="width:12%"><col style="width:22%"><col style="width:8%"><col style="width:14%"><col style="width:44%"></colgroup>
          <thead><tr><th>일자</th><th>대출처</th><th>구분</th><th>금액</th><th>적요</th></tr></thead>
          <tbody>{내역}</tbody></table></div>
      </div></div>"""


def C(v, 자릿수=0, 괄호=True):
    """표시 통화로 바꾼 금액 문자열. 음수는 (괄호) 빨간 글씨."""
    if v is None or pd.isna(v):
        return '-'
    t = f'{v * 배율 / 표시나눔:,.{자릿수}f}'
    return 음수표기(t) if 괄호 else t


def CK(v, 괄호=True):
    """천 단위 표기 (카드용) — 값만 돌려주고 단위는 K USD / K USD 로 따로 붙입니다."""
    if v is None or pd.isna(v):
        return '-'
    t = f'{v * 배율 / 표시나눔:,.0f}'
    if t in ('-0', '−0'):
        t = '0'
    return 음수표기(t) if 괄호 else t


def K(v):
    """그래프에 쓸 천 단위 숫자."""
    return None if v is None or pd.isna(v) else v * 배율 / 표시나눔


단위K = ('천원' if 통화 == 'KRW' and 표시나눔 == 1000 else (f'K {통화}' if 표시나눔 == 1000 else 통화))
환율문구 = '' if 환율 == 1.0 else f' · 환율 1 USD = {환율:.4f} {통화}'


def 차트제목(제목, 단위=None):
    """차트 위에 제목과 단위를 따로 그립니다 (Altair 제목은 글자가 잘려서 쓰지 않습니다)."""
    st.markdown(f"<div class='chart-head'><span class='t'>{제목}</span>"
                f"<span class='u'>(단위: {단위 or 단위K})</span></div>",
                unsafe_allow_html=True)


def 값라벨(차트, dy=-9, dx=0, align='center', 필드='금액'):
    """그래프에 천단위 쉼표 숫자를 붙입니다. 음수는 아래쪽에 표시합니다."""
    글꼴 = dict(fontSize=12, fontWeight=600, color=T['ink2'])
    양 = 차트.transform_filter(alt.datum[필드] >= 0).mark_text(
        dy=dy, dx=dx, align=align, **글꼴).encode(
        text=alt.Text(f'{필드}:Q', format=',.0f'), color=alt.value(T['ink2']))
    음 = 차트.transform_filter(alt.datum[필드] < 0).mark_text(
        dy=(-dy if dy < 0 else dy) + 4, dx=dx, align=align, **글꼴).encode(
        text=alt.Text(f'{필드}:Q', format=',.0f'), color=alt.value(T['ink2']))
    return 양 + 음


def X축(값들, 제목=None, 여유=0.20):
    """가로 막대 오른쪽 끝의 숫자가 잘리지 않도록 오른쪽 여백을 둡니다."""
    유효 = [float(v) for v in 값들 if v is not None and not pd.isna(v)]
    if not 유효:
        return alt.X('금액:Q', title=제목)
    lo, hi = min(유효 + [0.0]), max(유효 + [0.0])
    폭 = (hi - lo) or (abs(hi) or 1.0)
    return alt.X('금액:Q', title=제목,
                 scale=alt.Scale(domain=[lo - (폭 * 여유 if lo < 0 else 0), hi + 폭 * 여유],
                                 nice=False))


def Y축(값들, 제목=None, 여유=0.18):
    """막대 위·아래의 숫자가 축 글씨와 겹치지 않도록 위아래 여백을 넉넉히 둡니다."""
    유효 = [float(v) for v in 값들 if v is not None and not pd.isna(v)]
    if not 유효:
        return alt.Y('금액:Q', title=제목)
    lo, hi = min(유효 + [0.0]), max(유효 + [0.0])
    폭 = (hi - lo) or (abs(hi) or 1.0)
    return alt.Y('금액:Q', title=제목,
                 scale=alt.Scale(domain=[lo - 폭 * 여유, hi + 폭 * 여유], nice=False))

# ══════════════════════════════════════════════════════════════
# 3. 대시보드
# ══════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════
# 분석 엑셀은 「만들기」를 누를 때에만 만듭니다.
#   ※ 화면을 열 때마다 만들면 클릭 한 번에 5초 넘게 멈추고 메모리도 크게 써서,
#     웹사이트(Streamlit Cloud)가 앱을 강제로 내려 화면이 하얗게 보였습니다.
#     한 번 만든 파일은 원장이 바뀌기 전까지 다시 만들지 않습니다.
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner='분석 엑셀을 만드는 중입니다... (30초쯤 걸립니다)', max_entries=1)
def 분석엑셀캐시(열쇠):
    return 분석엑셀만들기(당해원장, 전년원장, P손익, P, 자금표, 현금잔고,
                          기초현금, 당해연도, 보고월, 매핑표)


def 엑셀열쇠():
    """원장이 바뀌면 값이 달라지는 「지문」 — 이 값이 같으면 엑셀을 다시 만들지 않습니다."""
    return (법인, 당해연도, 보고월, len(당해원장), len(전년원장),
            round(float(당해원장['금액'].abs().sum()), 2), round(float(기초현금), 2))


if 메뉴 in ('CTKonnect', '씨티케이커넥트'):
    제목칸, 내려받기칸 = st.columns([4, 1], gap='small')
    with 내려받기칸:
        if st.session_state.get('엑셀만들기'):
            try:
                엑셀바이트 = 분석엑셀캐시(엑셀열쇠())
                st.download_button(
                    '⤓  분석 엑셀 내려받기', data=엑셀바이트,
                    file_name=f'{법인}_실적분석_{당해연도}년{보고월}월누적_클로드작성.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    width='stretch',
                    help='지금 화면의 손익·현금흐름·미수채권과 처리본 원장을 담은 엑셀입니다')
            except Exception as e:
                st.session_state['엑셀만들기'] = False
                st.caption(f'엑셀 생성 실패: {e}')
        else:
            st.button('⤓  분석 엑셀 만들기', width='stretch', key='엑셀만들기버튼',
                      on_click=lambda: st.session_state.__setitem__('엑셀만들기', True),
                      help='누르면 손익·현금흐름·미수채권과 처리본 원장을 담은 엑셀을 만듭니다')
    with 제목칸:
        st.html(f"""<div class="wrap" translate="no" style="margin-bottom:0">
      <div class="page-head"><span class="t">실적보고 — {법인}</span>
        <span class="pill">{당해연도}년 1~{보고월}월 누적{' · 단위 천원' if 통화 == 'KRW' else ''}</span></div>
    </div>""")

    이전월 = 보고월 - 1 if 보고월 > 1 else None
    # ★KPI 카드는 1~보고월 누적★ (아래 표·그래프와 같은 기준)
    매출, 원가 = 당해누적['매출액'], 당해누적['매출원가']
    총이익, 영익 = 당해누적['매출총이익'], 당해누적['영업이익(손실)']
    # 원가 기준(재고수불부)이 기본 — 아래 손익계산서 카드의 라디오와 같은 기준으로 맞춥니다
    _수불 = (BS밑자료 or {}).get('수불원가') or {}
    KPI원가 = (bool(_수불.get('직접재료비'))
               and st.session_state.get('dash_pl_basis', 원가기준라벨) == 원가기준라벨
               and not str(st.session_state.get('dash_pl_kind', '')).startswith('활동')
               and (전년없음 or not str(st.session_state.get('dash_pl_year', '')).startswith(str(전년연도))))
    if KPI원가:
        _m0 = list(range(1, 보고월 + 1))
        _IS0 = 당해원장[당해원장['계정분류'].eq('IS')]
        _줄 = lambda n: [float(P.loc[n, m]) for m in _m0]
        C0 = 원가기준표(_IS0, _줄('매출액'), _줄('매출원가'), _줄('판매관리비'),
                        [-v for v in _줄('SR&ED 차감')], _m0, _수불)
        원가 = float(sum(C0['변동']) + sum(C0['간접']))
        총이익 = float(sum(C0['매출이익']))
        영익 = float(sum(C0['영익']))
    원가율 = (원가 / 매출 * 100) if 매출 else 0
    영익률 = (영익 / 매출 * 100) if 매출 else 0

    def 전년(항목):
        """누적끼리 비교해야 하니 전년 같은 기간(1~보고월) 누적을 돌려줍니다."""
        v = float(전년동기.get(항목, 0.0))
        return v if abs(v) > 0.005 else None

    누적표기 = f'1~{보고월}월 누적'
    if KPI원가:
        _sr0 = 당해원장[당해원장['계정영문'].astype(str).str.strip().str.lower().eq('sr&ed claim')
                       & 당해원장['월'].le(보고월)]

        def _칸(이름, 값, 아래, 강조=False):
            틀 = (f'border:2px solid {T["accent"]};background:{T["accentbg"]};' if 강조 else '')
            return (f'<div class="kpi-card mini" style="{틀}">'
                    f'<div class="kpi-label" style="{"color:" + T["accent"] + ";" if 강조 else ""}white-space:nowrap">{이름}</div>'
                    f'<div class="kpi-value" style="white-space:nowrap">{CK(값)} '
                    f'<span class="unit">{단위K}</span></div>'
                    f'<div class="kpi-delta">{아래}</div></div>')

        def _기호(t):
            return (f'<div style="display:flex;align-items:center;justify-content:center;font-size:16px;'
                    f'font-weight:800;color:{T["ink2"]}">{t}</div>')

        def _흐름(제목, 매, 변, 간, 판, 영후, 월들, 전년매, 머리말='', sr표시=True):
            """매출액 − 변동비 = 공헌이익 − 제조간접비 = 매출이익 − 판관비 = 영업이익 한 줄."""
            율 = lambda v: f'<span class="muted">매출 대비 {v / 매 * 100:,.1f}%</span>' if 매 else ''
            공 = 매 - 변
            sr = _sr0[_sr0['월'].isin(월들)]
            sr연 = sr.groupby(sr['Description'].astype(str).str.extract(r'(20\d\d)', expand=False)
                               .fillna('연도 미상'))['금액'].sum()
            sr글 = '<br>'.join(f'<span class="muted">SR&amp;ED {y}년분</span> {CK(abs(v))}'
                                for y, v in sorted(sr연.items()) if abs(v) >= 0.5)
            return (f'<div class="sub" style="margin:0 0 4px;font-size:12px"><b style="color:{T["ink"]}">{제목}</b> · 원가 기준(재고수불부) · '
                    f'매출액 − 변동비 = 공헌이익 − 제조간접비 = 매출이익 − 판관비 = 영업이익{머리말}</div>'
                    '<div style="display:grid;grid-template-columns:1fr 18px 1fr 18px 1.1fr 18px 1fr 18px 1fr 18px 1fr 18px 1fr;'
                    'gap:0;margin-bottom:10px;align-items:stretch">'
                    + _칸('매출액', 매, 증감HTML(매, 전년매, '전년동기')) + _기호('−')
                    + _칸('변동비 (매출원가)', 변, 율(변)) + _기호('=')
                    + _칸('공헌이익', 공, 율(공) + ' · ★ 목표 : 이익전환', 강조=True) + _기호('−')
                    + _칸('제조간접비', 간, 율(간)) + _기호('=')
                    + _칸('매출이익', 공 - 간, 율(공 - 간)) + _기호('−')
                    + _칸('판관비', 판, 율(판)) + _기호('=')
                    + (_칸('영업이익 (SR&amp;ED 차감 전)', 공 - 간 - 판,
                           f'<span class="muted">SR&amp;ED 차감 시</span> <b>{CK(영후)}</b>'
                           + (f'<br>{sr글}' if sr글 else '')) if sr표시 else
                       _칸('영업이익', 공 - 간 - 판, ''))
                    + '</div>')

        _i = 보고월 - 1
        _당월매 = float(_줄('매출액')[_i])
        _전년당월 = float(P전.loc['매출액', 보고월]) if 보고월 in P전.columns else None
        _가정월 = 보고월 in (C0.get('가정월') or [])
        kpi = (_흐름(f'당월 · {보고월}월', _당월매, float(C0['변동'][_i]), float(C0['간접'][_i]),
                     float(C0['판관'][_i]), float(C0['영익'][_i]), [보고월],
                     _전년당월 if _전년당월 else None,
                     (f' · <b style="color:{T["warn"]}">■ 가정 : {보고월}월 원가는 수불부 평균 원가율</b>' if _가정월 else ''), sr표시=False)
               + _흐름(누적표기, 매출, float(sum(C0['변동'])), float(sum(C0['간접'])),
                       float(sum(C0['판관'])), 영익, list(range(1, 보고월 + 1)), 전년('매출액')))
    else:
      kpi = f"""
    <div class="kpi-row" style="--n:4">
      <div class="kpi-card"><div class="kpi-label">매출액 <span class="calc">({누적표기})</span></div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{CK(매출)} <span class="unit">{단위K}</span></div>
        <div class="kpi-delta">{증감HTML(매출, 전년('매출액'), '전년동기')}</div></div>
      <div class="kpi-card"><div class="kpi-label">{'매출원가 (변동비 + 제조간접비)' if KPI원가 else '매출원가'} <span class="calc">({누적표기})</span></div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{CK(원가)} <span class="unit">{단위K}</span></div>
        <div class="kpi-delta">{'<span class="muted">원가 기준 (재고수불부)</span>' if KPI원가 else 증감HTML(원가, 전년('매출원가'), '전년동기')}</div></div>
      <div class="kpi-card"><div class="kpi-label">{'매출이익' if KPI원가 else '매출총이익'} <span class="calc">({누적표기})</span></div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{CK(총이익)} <span class="unit">{단위K}</span></div>
        <div class="kpi-delta"><span class="muted">원가율 {원가율:,.1f}%</span></div></div>
      <div class="kpi-card"><div class="kpi-label">영업이익(손실) <span class="calc">({누적표기})</span></div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{CK(영익)} <span class="unit">{단위K}</span></div>
        <div class="kpi-delta"><span class="muted">영업이익률 {영익률:,.1f}%</span></div></div>
    </div>"""

    months = list(range(1, 보고월 + 1))
    매출26 = [P.loc['매출액', m] for m in months]
    매출25 = [P전.loc['매출액', m] for m in months]
    # 맨 오른쪽 누적 막대는 두 해를 「같은 기간」으로 견줍니다 —
    # 전년 연간과 견주면 기간이 달라 매출이 줄어든 것처럼 보였습니다.
    누적26 = float(np.nansum(매출26))                    # 당해 1~보고월 누적
    누적25 = float(np.nansum(매출25))                    # 전년 같은 기간(1~보고월) 누적
    월최대 = max([abs(v) for v in 매출26 + 매출25] + [1])
    누적최대 = max(abs(누적26), abs(누적25), 1)

    def 막대(v25, v26, 라벨, 기준, 강조=False):
        cls = ' cum' if 강조 else ''
        return (f'<div class="bar-group{cls}"><div class="bar-pair">'
                f'<div class="bar-col"><div class="bar-val">{CK(v25)}</div>'
                f'<div class="bar prev" style="height:{abs(v25)/기준*118:.0f}px;"></div></div>'
                f'<div class="bar-col"><div class="bar-val">{CK(v26)}</div>'
                f'<div class="bar cur" style="height:{abs(v26)/기준*118:.0f}px;"></div></div>'
                f'</div><div class="bar-month">{라벨}</div></div>')

    bars = ''.join(막대(매출25[i], 매출26[i], f'{m}월', 월최대) for i, m in enumerate(months))
    bars += ('<div class="bar-sep"></div>'
             + 막대(누적25, 누적26, f'1~{보고월}월 누적', 누적최대, True))

    # ── 월별 매출 추이 : 월별은 꺾은선, 1~보고월 누적은 오른쪽 막대 두 개 (2026-10-27 요청)
    #    st.html 은 <svg> 를 지우므로 그림(data URI)으로 넣습니다
    def _선그래프():
        import base64 as _b64
        _백만 = 법인 == '씨티케이커넥트'
        CK = (lambda v, 괄호=True: '-' if v is None or pd.isna(v) else
              음수표기(f'{v / 1_000_000:,.0f}')) if _백만 else globals()['CK']
        W, H, 위, 아래, 왼, 오른 = 1000, 175, 30, 36, 20, 170   # 높이 30% 줄임 (250 → 175, 2026-10-29)
        폭 = W - 왼 - 오른
        n = max(len(months), 1)
        최대 = max([*매출25, *매출26, 1]) * 1.12
        x = lambda i: 왼 + 폭 * (i + 0.5) / n
        y = lambda v: 위 + (H - 위 - 아래) * (1 - max(v, 0) / 최대)
        g = ''
        for k in range(1, 4):            # 옅은 가로선
            yy = 위 + (H - 위 - 아래) * k / 4
            g += f'<line x1="{왼}" x2="{왼 + 폭}" y1="{yy:.1f}" y2="{yy:.1f}" stroke="#e5e7eb" stroke-width="1"/>'
        for 값들, 색, 당해 in (((매출26, T['accent'], True),) if 전년없음 else ((매출25, SLATE, False), (매출26, T['accent'], True))):
            pts = ' '.join(f'{x(i):.1f},{y(v):.1f}' for i, v in enumerate(값들))
            g += f'<polyline points="{pts}" fill="none" stroke="{색}" stroke-width="3" stroke-linejoin="round"/>'
            for i, v in enumerate(값들):
                g += f'<circle cx="{x(i):.1f}" cy="{y(v):.1f}" r="4.5" fill="#fff" stroke="{색}" stroke-width="2.5"/>'
                다른 = (매출25 if 당해 else 매출26)[i]
                ty = y(v) - 11 if v >= 다른 else y(v) + 20
                g += (f'<text x="{x(i):.1f}" y="{ty:.1f}" text-anchor="middle" font-size="14" '
                      f'font-weight="700" fill="{색 if 당해 else "#4b5563"}" stroke="#fff" stroke-width="5" '
                      f'stroke-linejoin="round" paint-order="stroke">{CK(v)}</text>')
        for i, m in enumerate(months):
            g += f'<text x="{x(i):.1f}" y="{H - 12}" text-anchor="middle" font-size="14" fill="#6b7280">{m}월</text>'
        구분x = W - 오른 + 18
        g += f'<line x1="{구분x}" x2="{구분x}" y1="{위 - 10}" y2="{H - 30}" stroke="#d1d5db"/>'
        누최대 = max(누적25, 누적26, 1) * 1.15
        for j, (v, 색) in enumerate((((누적26, T['accent']),) if 전년없음 else ((누적25, SLATE), (누적26, T['accent'])))):
            bx = 구분x + 40 + j * 50
            top = 위 + (H - 위 - 아래) * (1 - max(v, 0) / 누최대)
            g += f'<rect x="{bx}" y="{top:.1f}" width="26" height="{H - 아래 - top:.1f}" rx="4" fill="{색}"/>'
            g += (f'<text x="{bx + 13}" y="{top - 7:.1f}" text-anchor="middle" font-size="14" '
                  f'font-weight="700" fill="#111827">{CK(v)}</text>')
        g += (f'<text x="{구분x + 78}" y="{H - 12}" text-anchor="middle" font-size="14" '
              f'font-weight="700" fill="#111827">1~{보고월}월 누적</text>')
        svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" '
               f'font-family="Malgun Gothic, Pretendard, sans-serif">{g}</svg>')
        return (f'<img alt="월별 매출 추이" style="width:100%;height:auto;display:block;margin-top:6px" '
                f'src="data:image/svg+xml;base64,{_b64.b64encode(svg.encode("utf-8")).decode()}">')
    try:
        bars_line = _선그래프()
    except Exception:                      # 그림이 안 되면 예전 막대로
        bars_line = f'<div class="bars">{bars}</div>'

    # ── 매출액 구성 (손익 기준) ─────────────────────────────
    # KPI 카드·손익표와 같게 1~보고월 누적으로 봅니다
    매출행 = 당해원장[당해원장['월'].le(보고월) & (당해원장['분류'] == '매출')]
    mix = 매출행.groupby('계정과목')['금액'].sum().sort_values(ascending=False)
    mix_html = ''
    if mix.sum():
        for 이름, 값 in mix.items():
            비율 = 값 / mix.sum() * 100
            mix_html += (f'<div class="mix-item"><div class="top"><span>{이름}</span>'
                         f'<span>{CK(값)} <span class="muted">{단위K} ({비율:,.0f}%)</span></span></div>'
                         f'<div class="mix-track">'
                         f'<div class="mix-fill" style="width:{max(0, min(100, 비율)):.0f}%;"></div></div></div>')
        mix_html += (f'<div class="mix-total">매출액 합계<span>{CK(mix.sum())} '
                     f'<span class="muted">{단위K}</span></span></div>')
    else:
        mix_html = '<div class="muted">해당 기간 매출 데이터가 없습니다.</div>'

    # ══ 손익계산서 표 ══════════════════════════════════════
    #    카드 위 라디오 : 연도(당해/전년) · 판관비 구분(손익계산서 비용/활동성 비용) · 손익 기준(퀵북/원가)
    당해보기 = 전년없음 or not str(st.session_state.get('dash_pl_year', '')).startswith(str(전년연도))
    활동기준 = str(st.session_state.get('dash_pl_kind', '')).startswith('활동')
    원장X = 당해원장 if 당해보기 else 전년원장
    PX = P if 당해보기 else P전                      # 활동분류 기준 표 (합계 줄은 손익 기준과 같음)
    보고월X = 보고월 if 당해보기 else int(전년원장['월'].max() or 12)
    monthsX = list(range(1, 보고월X + 1))
    연도X = 당해연도 if 당해보기 else 전년연도
    IS = 원장X[원장X['계정분류'].eq('IS')]

    def 월별(mask, 부호=1):
        s = IS[mask].groupby('월')['금액'].sum().reindex(monthsX, fill_value=0.0) * 부호
        return [float(s.get(m, 0.0)) for m in monthsX]

    # ★실적집계 화면과 같은 표(P)에서 그대로 가져옵니다 — 두 화면이 어긋날 수 없습니다★
    def 줄(이름):
        return [float(PX.loc[이름, m]) for m in monthsX]

    매출L, 원가L, 총이익L = 줄('매출액'), 줄('매출원가'), 줄('매출총이익')
    총액자료 = (BS밑자료 or {}).get('총액')
    총액보기 = bool(총액자료) and 당해보기 and st.session_state.get('dash_pl_gross') == '총액매출'
    판관L, 판관순L = 줄('판매관리비'), 줄('판관비 (SR&ED 반영)')
    SREDL = [-v for v in 줄('SR&ED 차감')]
    영익L = 줄('영업이익(손실)')
    # 영업이익 아래 — 이자·외환 같은 영업 밖 손익과 세전이익.
    #   이자비용이 영업적자보다 큰 달이 있어, 영업이익까지만 보면 실제 적자가 가려집니다.
    영업외L = 월별(~IS['분류'].isin(['매출', '매출원가', '판관비']))
    세전L = [영익L[i] + 영업외L[i] for i in range(len(monthsX))]

    리스상각 = 당해원장['계정영문'].astype(str).str.strip().str.lower().eq('depreciation - rou - lease')
    상각L = 월별(IS['계정과목'].isin(['감가상각비', '무형자산상각비']), -1)
    리스L = 월별(IS['계정과목'].isin(['감가상각비', '무형자산상각비']) & 리스상각, -1)

    판관계정 = (IS[IS['분류'].eq('판관비')].groupby('계정과목')['금액'].sum().mul(-1)
                .sort_values(ascending=False))
    상위4 = list(판관계정.index[:4])
    세부, 세부들여 = [], {}
    if 활동기준:
        # 활동분류(보고항목) 기준 — 월별 실적집계 화면과 같은 줄 (상각비 아래 감가상각비 · 리스상각비)
        수준표X = PX.attrs['수준']
        줄들 = list(PX.index)
        시작, 끝 = 줄들.index('판매관리비') + 1, 줄들.index('SR&ED 차감')
        for 이름 in 줄들[시작:끝]:
            세부.append((이름, 줄(이름)))
            세부들여[이름] = max(1, int(수준표X.get(이름, 1)))
        활동항목들 = [n for n, _ in 세부 if 세부들여[n] == 1]
    else:
        for 이름 in 상위4:
            세부.append((이름, 월별(IS['분류'].eq('판관비') & IS['계정과목'].eq(이름), -1)))
        세부.append(('기타', [판관L[i] - sum(v[1][i] for v in 세부) for i in range(len(monthsX))]))
        활동항목들 = []
    세부.append(('SR&ED 차감', [-v for v in SREDL]))

    def 이름칸(이름, 산식=''):
        """산식은 어떻게 계산했는지 덧붙이는 설명이라 () 안에 한 포인트 작게 씁니다."""
        return 이름 + (f' <span class="calc">({산식})</span>' if 산식 else '')

    예상월들 = list(range(보고월X + 1, 13))        # 8월(예상) … 12월(예상)

    # ── 8~12월 예상 — 현금흐름 표와 똑같은 가정(1~보고월X 실적 월평균)을 손익에도 깝니다.
    #    · 매출액   = 실적 월평균
    #    · 매출원가 = 실적 월평균 (= 예상 매출액 × 실적 원가율, 같은 값)
    #    · 판관비   = 실적 월평균. SR&ED 차감은 한 번 받고 마는 돈이라 0 으로 둡니다
    #                 (상각비는 현금이 아닐 뿐 비용은 비용이라 손익에는 그대로 둡니다)
    #    · 영업외손익(이자·환율)은 차입 조건에 달린 일이라 가정하지 않습니다
    #      → 그래서 영업외손익 · 세전이익의 예상 칸은 비워 둡니다.
    def 월평(값들):
        return float(sum(값들)) / max(보고월X, 1)

    예상매출, 예상원가 = 월평(매출L), 월평(원가L)
    예상판관 = 월평(판관L)                       # SR&ED 반영 전 (= 일회성 제외)
    예상총이익 = 예상매출 - 예상원가
    예상영익 = 예상총이익 - 예상판관
    예상세부 = {이름: (0.0 if 이름 == 'SR&ED 차감' else 월평(값들))
                for 이름, 값들 in 세부}

    def 행(이름, 값들, 굵게=False, 들여=0, 부호표시=False, 색='', 산식='', 예상=None, 강조=False,
           숨김=False, 토글=False):
        cls = (' class="total hl"' if 강조 else ' class="total"' if 굵게
               else (' class="sub sga"' if 숨김 else ' class="sub"' if 들여 else ''))
        def f(v):
            if abs(v * 배율 / 표시나눔) < 0.5:   # 반올림하면 0 이 되는 금액은 '-' 로
                return '-'
            if 부호표시:
                return ('+' if v > 0 else '−') + CK(abs(v))
            return CK(v)
        cells = ''.join(f'<td class="{색}">{f(v)}</td>' for v in 값들)
        # 앞으로 올 달(예상) 칸 — 가정을 깐 줄만 채우고, 나머지는 비워 둡니다
        칸값 = '-' if 예상 is None else f(예상)
        cells += ''.join(f'<td class="fc {색}">{칸값}</td>' for _ in 예상월들)
        총 = sum(값들)
        return (f'<tr{cls}><td class="name" style="white-space:nowrap;overflow:hidden;'
                f'text-overflow:ellipsis" title="{re.sub(r"<[^>]+>", "", 이름칸(이름, 산식))}">'
                f'{"&nbsp;&nbsp;" * max(들여 - 1, 0)}'
                + (f'<label class="tog" for="pl-sga">{이름칸(이름, 산식)}<span class="caret">▸</span></label>'
                   if 토글 else 이름칸(이름, 산식))
                + f'</td>{cells}'
                f'<td class="{색}"><b>{f(총)}</b></td></tr>')

    # 실적 뒤 ~ 연말까지 얼마를 깔았는지, 줄 이름 옆에 한 줄로 적어 둡니다 (30자 안쪽)
    달수예 = len(예상월들)
    원가율 = (sum(원가L) / sum(매출L) * 100) if sum(매출L) else 0.0
    근거매출 = (f'{예상월들[0]}~12월 예상 {CK(예상매출 * 달수예)} = 월 {CK(예상매출)} × {달수예}'
                if 예상월들 else '')
    근거원가 = (f'{예상월들[0]}~12월 예상 {CK(예상원가 * 달수예)} = 실적 원가율 {원가율:,.1f}%'
                if 예상월들 else '')
    근거판관 = (f'{예상월들[0]}~12월 예상 {CK(예상판관 * 달수예)} = 실적 월평균 × {달수예}'
                if 예상월들 else '')

    # ── 손익계산서 기준 : 원가 기준(재고수불부, 기본) / 퀵북 원장 — 카드 위 라디오에서 고릅니다
    수불원가 = (BS밑자료 or {}).get('수불원가') or {}
    원가가능 = bool(수불원가.get('직접재료비')) and 당해보기 and not 활동기준
    원가기준 = 원가가능 and st.session_state.get('dash_pl_basis', 원가기준라벨) == 원가기준라벨
    원가설명 = ''
    if 원가기준:
        C = 원가기준표(IS, 매출L, 원가L, 판관L, SREDL, monthsX, 수불원가)
        재료L, 외주L, 노무L, 간접L, 변동L = C['재료'], C['외주'], C['노무'], C['간접'], C['변동']
        매출이익L, 판관원가L, 세부원가, 판관순원가L, 영익원가L = C['매출이익'], C['판관'], C['세부'], C['판관순'], C['영익']
        가정월 = C['가정월']
        세전원가L = [영익원가L[i] + 영업외L[i] for i in range(len(monthsX))]
        차이L = [영익원가L[i] - 영익L[i] for i in range(len(monthsX))]
        예상재료, 예상외주, 예상노무, 예상간접 = 월평(재료L), 월평(외주L), 월평(노무L), 월평(간접L)
        예상변동 = 예상재료 + 예상외주 + 예상노무
        예상매출이익 = 예상매출 - 예상변동 - 예상간접
        예상판관원가 = 월평(판관원가L)
        예상영익원가 = 예상매출이익 - 예상판관원가
        예상세부원가 = {이름: (0.0 if 이름 == 'SR&ED 차감' else 월평(값들)) for 이름, 값들 in 세부원가}
        가정글 = (f'{"·".join(str(m) for m in 가정월)}월은 가정 — 수불부가 있는 '
                  f'{C["수불월"][0]}~{C["수불월"][-1]}월 평균 원가율(재료비 {C["재료율"]*100:.1f}% · 외주 {C["외주율"]*100:.1f}%) × 그 달 매출'
                  if 가정월 and C['수불월'] else '')
        rows = 행('매출액', 매출L, True, 예상=예상매출)
        rows += 행('변동비', 변동L, True, 예상=예상변동)
        rows += 행('직접재료비 (원부자재)', 재료L, 들여=1, 예상=예상재료)
        rows += 행('외주가공비 (외주생산)', 외주L, 들여=1, 예상=예상외주)
        rows += 행('직접노무비 (생산인력)', 노무L, 들여=1, 예상=예상노무)
        공헌L = [매출L[i] - 변동L[i] for i in range(len(monthsX))]
        rows += 행('공헌이익', 공헌L, True, 예상=예상매출 - 예상변동, 산식='매출액 − 직접원가', 강조=True)
        rows += 행('고정비', 간접L, True, 예상=예상간접)
        rows += 행('제조간접비 (감가상각비 등)', 간접L, 들여=1, 예상=예상간접)
        rows += 행('매출이익', 매출이익L, True, 예상=예상매출이익)
        # 판관비 합계는 SR&ED 차감 전, SR&ED 는 따로 한 줄 (2026-10-27 요청)
        rows += 행('판관비 합계', 판관원가L, True, 예상=예상판관원가, 토글=True)
        for 이름, 값들 in 세부원가:
            if 이름 != 'SR&ED 차감':
                rows += 행(이름, 값들, 들여=1, 예상=예상세부원가.get(이름), 숨김=True)
        rows += 행('영업이익(손실)', 영익원가L, True, 예상=예상영익원가)
        rows += 행('영업외손익', 영업외L, 산식='이자 · 외환 등')
        rows += 행('세전이익(손실)', 세전원가L, True)
        가정짧게 = (f'{"·".join(str(m) for m in 가정월)}월 원가 = {C["수불월"][0]}~{C["수불월"][-1]}월 평균 원가율'
                    f'(재료 {C["재료율"]*100:.1f}% · 외주 {C["외주율"]*100:.1f}%) × 매출'
                    if 가정월 and C['수불월'] else '')
        원가설명 = ('<li>직접재료비 · 외주가공비 = 재고수불부 / 직접노무비 · 제조간접비 = 원장</li>'
                    + (f'<li class="red">■ 가정 : {가정짧게}</li>' if 가정짧게 else ''))
    else:
        _코드 = IS['계정영문'].astype(str).str.strip()
        중개L, IORL, 상품L = (월별(_코드.str.startswith(k)) for k in ('51103', '51104', '51101'))
        상품원가L = 월별(_코드.str.startswith('52101'), -1)
        if 총액보기:
            _총 = 총액자료
            매입L = [float(_총['중개매입'].get(m, 0.0)) for m in monthsX]
            대납L = [float(_총['IOR대납'].get(m, 0.0)) for m in monthsX]
            rows = 행('매출액 (총액 · 거래액)', [매출L[i] + 매입L[i] + 대납L[i] for i in range(len(monthsX))], True)
            rows += 행('상품수수료매출 (중개)', [중개L[i] + 매입L[i] for i in range(len(monthsX))], 들여=1)
            rows += 행('IOR매출', [IORL[i] + 대납L[i] for i in range(len(monthsX))], 들여=1)
            rows += 행('상품매출', 상품L, 들여=1)
            rows += 행('매출원가 (총액)', [원가L[i] + 매입L[i] + 대납L[i] for i in range(len(monthsX))], True)
            rows += 행('중개 매입대금', 매입L, 들여=1)
            rows += 행('IOR 대납 물류 · 통관비', 대납L, 들여=1)
            rows += 행('상품매출원가', 원가L, 들여=1)
        else:
            rows = 행('매출액', 매출L, True, 예상=예상매출)
            if 법인 == '씨티케이커넥트':        # 한국 법인 : 계정과목별 매출 줄
                for _n in ('제품매출', '상품매출', '용역매출', '기타매출'):
                    _v = 월별(IS['계정과목'].astype(str).eq(_n))
                    if any(abs(x) > 0.5 for x in _v):
                        rows += 행(_n, _v, 들여=1)
            else:
                rows += 행('상품수수료매출 (중개)', 중개L, 들여=1)
                rows += 행('IOR매출', IORL, 들여=1)
                rows += 행('상품매출', 상품L, 들여=1)
            rows += 행('매출원가', 원가L, 예상=예상원가)
        rows += 행('매출총이익', 총이익L, True, 예상=예상총이익)
        # 판관비 합계는 SR&ED 차감 전, SR&ED 는 따로 한 줄 (2026-10-27 요청)
        rows += 행('판관비 합계', 판관L, True, 예상=예상판관, 토글=True)
        for 이름, 값들 in 세부:
            if 이름 != 'SR&ED 차감':
                rows += 행(이름, 값들, 들여=세부들여.get(이름, 1), 예상=예상세부.get(이름), 숨김=True)
        rows += 행('영업이익(손실)', 영익L, True, 예상=예상영익)
        rows += 행('영업외손익', 영업외L, 산식='이자 · 외환 등')
        rows += 행('세전이익(손실)', 세전L, True)

    # 열너비를 %로 잡아 화면 크기에 맞춰 표가 같이 줄었다 늘었다 하게 합니다.
    # 손익계산서와 현금흐름이 똑같은 %와 똑같은 최대너비를 쓰므로, 화면이 크든 작든
    # 두 표의 월별 칸은 항상 같은 자리에서 시작합니다 (위아래 라인이 맞습니다).
    칸수 = len(monthsX) + len(예상월들) + 1        # 월 칸들 + 예상 칸들 + 「누적」 칸
    이름칸, 숫자칸 = 200, 98                    # CTKonnect : USD 그대로라 자릿수가 커서 숫자칸을 넓힘 (캐나다 285 · 78)
    이름비율 = 이름칸 / (이름칸 + 숫자칸 * 칸수) * 100
    숫자비율 = (100 - 이름비율) / 칸수
    표너비 = 이름칸 + 숫자칸 * 칸수               # 이보다 커지면 보기 나쁘니 최대너비로 씁니다
    콜그룹 = (f'<colgroup><col style="width:{이름비율:.4f}%">'
              + f'<col style="width:{숫자비율:.4f}%">' * 칸수 + '</colgroup>')
    표스타일 = f'style="width:100%; min-width:{표너비 * 0.92:.0f}px; max-width:{표너비}px; table-layout:fixed;"'

    # ══ 오른쪽 세부내역 : 손익계산서 항목을 고르면 거래처별 누적 내역 ══════
    누적IS = IS[IS['월'].le(보고월X)]

    def 항목조건_손익(이름):
        d = 누적IS
        if 이름 == '매출액':
            return d['분류'].eq('매출'), 1
        if 이름 == '매출원가':
            return d['분류'].eq('매출원가'), -1
        if 이름 == '판관비 합계':
            return d['분류'].eq('판관비'), -1
        if 이름 == '기타':
            return d['분류'].eq('판관비') & ~d['계정과목'].isin(상위4), -1
        if 활동기준 and '보고항목' in d.columns:
            return d['분류'].eq('판관비') & d['보고항목'].eq(이름), -1
        return d['분류'].eq('판관비') & d['계정과목'].eq(이름), -1

    # 「판관비 합계」는 아래 상위4 + 기타로 쪼개지는 상위 줄이라 목록에서 뺍니다 (최하위 줄만)
    드릴목록_손익 = ['매출액', '매출원가'] + (활동항목들 if 활동기준 else 상위4 + ['기타'])

    def CK1(v):
        """계산 근거는 덧셈·뺄셈 식이라 괄호 대신 부호를 두고 색만 붉게 씁니다."""
        t = f'{v * 배율 / 표시나눔:,.{1 if 표시나눔 > 1 and 통화 != "KRW" else 0}f}'
        return f'<span class="neg">{t}</span>' if t[:1] in ('-', '−') else t

    # 영업외손익이 무엇으로 이루어져 있는지 — 이자가 영업적자보다 큰 달이 있어 따로 적습니다
    영업외행 = IS[~IS['분류'].isin(['매출', '매출원가', '판관비'])]
    영업외계정 = 영업외행.groupby('계정과목')['금액'].sum().sort_values()
    영업외조각 = ' · '.join(f'{이름} {CK1(v)}' for 이름, v in 영업외계정.items()
                            if abs(v * 배율 / 표시나눔) >= 0.05)
    _영익표 = 영익원가L if 원가기준 else 영익L
    _세전표 = 세전원가L if 원가기준 else 세전L
    영업외설명 = (f'<li>영업외손익 : {영업외조각 or "해당 없음"} → 세전이익 '
                  f'<b>{CK1(sum(_세전표))}</b> (영업이익 {CK1(sum(_영익표))} + 영업외 {CK1(sum(영업외L))})</li>')

    if not 당해보기:          # 아래 현금흐름 표의 가정은 언제나 당해 실적으로 깝니다
        _m당 = list(range(1, 보고월 + 1))
        매출L = [float(P.loc['매출액', m]) for m in _m당]
        원가L = [float(P.loc['매출원가', m]) for m in _m당]
        판관L = [float(P.loc['판매관리비', m]) for m in _m당]
        _IS당 = 당해원장[당해원장['계정분류'].eq('IS')]
        _상각 = (_IS당[_IS당['계정과목'].isin(['감가상각비', '무형자산상각비'])]
                 .groupby('월')['금액'].sum().reindex(_m당, fill_value=0.0) * -1)
        상각L = [float(_상각.get(m, 0.0)) for m in _m당]

    # ══ 자금 변동 (Statement of Cash Flows) ══════════════════
    자금html = ''
    if 자금표 is not None:
        기초 = [float(현금잔고.get(m - 1, 기초현금)) if m > 1 else float(기초현금)
                for m in months]
        기말 = [float(현금잔고.get(m, 0.0)) for m in months]

        # ── 8월(예상) … 12월(예상) — 「26년 자금」 시트의 2번(출금)·3번(입금)에서
        #   · 출금·회수는 원장 실적에서 뽑은 가정으로 화면이 직접 계산합니다(예상현금모형).
        #   · 「26년 자금」 3번 표(담당자가 손으로 채우는 칸)에 값이 있으면 그 줄·그 달은
        #     담당자 값으로 갈아 끼웁니다 — 담당자가 넣은 숫자가 언제나 우선입니다.
        자금예상월 = list(range(보고월 + 1, 19))          # 다음 해 6월(18)까지 (2026-11-03 · 모든 법인 같게)
        예상, 가정 = 예상현금모형(매출L, 원가L, 판관L, 상각L, 자금표, 보고월, 끝월=18)

        def 달이름(m):
            return f'{m}월' if m <= 12 else f'{m - 12}월'

        def 달표(m):
            return f'{m}월' if m <= 12 else f"{(당해연도 + 1) % 100}년 {m - 12}월"
        시트예상 = 예상자금읽기(데이터바이트) or {}
        담당자칸 = (시트예상 or {}).get('3', {})
        for 줄, 칸 in 담당자칸.items():
            for m, v in 칸.items():
                if m <= 보고월:
                    continue
                예상.setdefault(줄, {})
                예상[줄][m] = v                 # 더하지 않고 대체합니다

        def 예(키, m):
            return float(예상.get(키, {}).get(m, 0.0))

        def 묶음(키, m):
            """소계 줄은 아래 줄을 더해서 냅니다."""
            if 키 == '지급계':
                return 예('AP', m) + 예('PAY', m) + 예('OTHOP', m)
            if 키 == '영업':
                return 예('AR', m) + 예('ETCIN', m) + 묶음('지급계', m)
            if 키 == '투자':
                return 예('CAPEX', m) + 예('INV', m)
            if 키 == '재무':
                return 예('EQ', m) + 예('차입증감', m) + 예('리스', m)
            if 키 == '순증감':
                return (묶음('영업', m) + 묶음('투자', m) + 묶음('재무', m) + 예('FX', m))
            return 예(키, m)

        def ㅈ(키):
            v = [float(자금표.loc[키, m]) for m in months]
            앞 = [float(np.nansum(v))]           # 「실적 누적」(1~보고월) — 맨 앞 칸
            return 앞 + v + [묶음(키, m) for m in 자금예상월]

        # 예상 달의 기말은 마지막 실적 잔액에서 이어 붙입니다
        예상기초, 예상기말, 앞잔액 = [], [], (기말[-1] if 기말 else 기초현금)
        for m in 자금예상월:
            예상기초.append(앞잔액)
            앞잔액 = 앞잔액 + 묶음('순증감', m)
            예상기말.append(앞잔액)

        자금줄 = [
            ('기초 현금', 0, [기초[0]] + 기초 + 예상기초, True),
            ('영업활동', 0, ㅈ('영업'), True),
            ('매출채권 회수', 1, ㅈ('AR'), False),
            ('기타입금', 1, ㅈ('ETCIN'), False),
            ('매입채무', 1, ㅈ('AP'), False),
            ('급여', 1, ㅈ('PAY'), False),
            ('기타 비용', 1, ㅈ('OTHOP'), False),
            ('투자활동', 0, ㅈ('투자'), True),
            ('설비 투자', 1, ㅈ('CAPEX'), False),
            ('금융자산 투자', 1, ㅈ('INV'), False),
            ('재무활동', 0, ㅈ('재무'), True),
            ('유상증자', 1, ㅈ('EQ'), False),
            ('차입증감', 1, ㅈ('차입증감'), False),
            ('리스료 지급', 1, ㅈ('리스'), False),
            ('환율변동 효과', 0, ㅈ('FX'), False),
            ('순증감', 0, ㅈ('순증감'), True),
            ('기말 현금', 0, [기말[-1]] + 기말 + 예상기말, True),      # 실적 누적 칸 = 실제 {보고월}월 말 잔액
        ]
        # 칸을 좁혀 한 화면에 — 구분 · 실적 누적 · 실적 달 · 예상 달 (2026-11-03)
        자금칸수 = 1 + len(months) + len(자금예상월)
        자금이름칸, 자금숫자칸 = 86, 58
        자금표너비 = 자금이름칸 + 자금숫자칸 * 자금칸수
        자금이름비율 = 자금이름칸 / 자금표너비 * 100
        자금숫자비율 = (100 - 자금이름비율) / 자금칸수
        자금콜그룹 = (f'<colgroup><col style="width:{자금이름비율:.4f}%">'
                      + f'<col style="width:{자금숫자비율:.4f}%">' * 자금칸수 + '</colgroup>')
        자금표스타일 = 'class="lined cfcomp" style="width:100%; table-layout:fixed;"'
        자금최소폭 = 자금표너비              # 이보다 좁으면 그림과 표를 같이 옆으로 밀어 봅니다

        자금rows = ''
        for 국문, 깊이, 값들, 굵게 in 자금줄:
            cls = ' class="total"' if 굵게 else (' class="sub sub2"' if 깊이 == 2
                                                 else (' class="sub"' if 깊이 else ''))
            예상칸 = ' class="fc"'
            칸 = ''.join(
                '<td' + (' class="cum"' if i == 0 else (예상칸 if i > len(months) else ''))
                + '>' + ('-' if abs(v * 배율 / 표시나눔) < 0.5 else CK(v)) + '</td>'
                for i, v in enumerate(값들))
            자금rows += f'<tr{cls}><td class="name">{국문}</td>{칸}</tr>'

        # ── 입금 · 변동비 · 고정비 (2026-11-03) — 위 「영업활동」 을 다시 나눈 것 · 합이 영업활동과 같음
        #    변동비 = 매출에 따라 늘고 주는 돈 : 매입대금(매입채무 · 선급금, 그리고 매입처에 낸 미지급금) · 판매 · 물류비
        #    고정비 = 매출이 없어도 매달 나가는 돈 : 급여 · 임차료 · 보험 · 통신 · 자문 · 수수료 · 이자 등 나머지
        고변html = ''
        if 자금상세 is not None and '계정과목' in 자금상세.columns:
            _상 = 자금상세.copy()
            _상['과목'] = _상['계정과목'].astype(str)
            _상['거래처'] = _상['Name'].fillna('').astype(str).str.strip()
            _변동과목 = {'운반비', '판매수수료', '외주용역수수료', '수출제비용', '판매촉진비', '광고선전비',
                        '포장비', '상품매출원가', '제품매출원가', '용역매출원가'}
            _선급매입 = 법인 == 'CTKonnect'          # CTKonnect 선급금 = 중개 매입대금 선지급 · 씨티케이커넥트는 자문료 선급도 섞임
            _매입과목 = ({'매입채무', '상품', '제품', '원재료', '재공품', '미착품'} | _변동과목
                         | ({'선급금'} if _선급매입 else set()))
            _원 = 당해원장
            _매입처 = set(_원[_원['계정과목'].astype(str).isin(_매입과목)]['Name'].dropna().astype(str).str.strip()) - {''}

            def _나눔(r):
                if r['구분'] in ('AR', 'ETCIN'):
                    return '입금'
                if r['구분'] == 'AP':
                    if r['과목'] == '미지급비용':
                        return '고정-급여' if ('임직원' in r['거래처'] or '급여' in r['거래처']) else '고정-기타'
                    if r['과목'] == '매입채무' or (r['과목'] == '선급금' and _선급매입) or r['거래처'] in _매입처:
                        return '변동-매입'
                    return '고정-기타'
                if r['구분'] == 'PAY':
                    return '고정-급여'
                if r['구분'] == 'OTHOP':
                    return '변동-판매' if r['과목'] in _변동과목 else '고정-기타'
                return None
            _상['나눔'] = _상.apply(_나눔, axis=1)
            _피 = (_상.dropna(subset=['나눔']).pivot_table(index='나눔', columns='월', values='영향', aggfunc='sum')
                   .reindex(columns=months).fillna(0.0))

            def _월(k):
                return [float(_피.loc[k, m]) if k in _피.index else 0.0 for m in months]
            입금L = [a + b for a, b in zip(_월('입금'), [0.0] * len(months))]
            변매L, 변판L = _월('변동-매입'), _월('변동-판매')
            고급L, 고기L = _월('고정-급여'), _월('고정-기타')
            변동L = [a + b for a, b in zip(변매L, 변판L)]
            고정L = [a + b for a, b in zip(고급L, 고기L)]
            순L = [a + b + c for a, b, c in zip(입금L, 변동L, 고정L)]
            영업실적 = [float(자금표.loc['영업', m]) for m in months]
            맞음 = all(abs(a - b) < 1 for a, b in zip(순L, 영업실적))
            월평균고정 = -sum(고정L) / max(len(months), 1)
            월평균변동 = -sum(변동L) / max(len(months), 1)
            월평균입금 = sum(입금L) / max(len(months), 1)
            말현금 = 기말[-1] if 기말 else 기초현금
            버팀 = (말현금 / 월평균고정) if 월평균고정 > 0.5 else None

            def _줄(이름, 값들, 굵게=False, 깊이=0, 중간=False):
                cls = ' class="total"' if 굵게 else (' class="sub"' if 깊이 else (' class="mid"' if 중간 else ''))
                칸 = ''.join('<td' + (' class="cum"' if i == 0 else '') + '>'
                             + ('-' if abs(v * 배율 / 표시나눔) < 0.5 else CK(v)) + '</td>'
                             for i, v in enumerate([sum(값들)] + 값들 + [sum(값들) / max(len(months), 1)]))
                return f'<tr{cls}><td class="name">{이름}</td>{칸}</tr>'
            _행 = (_줄('입금 (채권 · 기타)', 입금L, 중간=True)
                   + _줄('변동비', 변동L, 중간=True)
                   + _줄('매입대금 (매입 · 선급금 등)', 변매L, 깊이=1)
                   + _줄('판매 · 물류비 (운반 · 수수료 등)', 변판L, 깊이=1)
                   + _줄('고정비', 고정L, 중간=True)
                   + _줄('급여 · 인건비', 고급L, 깊이=1)
                   + _줄('기타 (임차 · 수수료 · 이자 등)', 고기L, 깊이=1)
                   + _줄('영업 현금 순증감 (= 위 표 「영업활동」)', 순L, True))
            _너비 = 자금이름비율 * 2.6 + 자금숫자비율 * (len(months) + 2)
            _콜 = (f'<colgroup><col style="width:{자금이름비율 * 2.6 / _너비 * 100:.4f}%">'
                   + f'<col style="width:{자금숫자비율 / _너비 * 100:.4f}%">' * (len(months) + 2) + '</colgroup>')
            고변html = (
                f'<h3 style="margin-top:18px">입금 · 변동비 · 고정비 <span class="unitbadge">단위 {단위K}</span></h3>'
                f'<div class="sub">위 「영업활동」 을 다시 나눈 표 · 실적(1~{보고월}월)만 · '
                + ('<b>합계가 위 영업활동과 같습니다</b>' if 맞음 else
                   f'<span style="color:{T["warn"]};font-weight:700">⚠ 위 영업활동과 다른 달이 있습니다 — 확인 필요</span>')
                + f' · 투자 · 재무(차입 · 증자 · 금융자산)는 회사가 정하는 돈이라 뺐습니다</div>'
                f'<div class="cfsum">한 달 평균 — 입금 <b>{CK(월평균입금)}</b> · 변동비 <b>{CK(월평균변동)}</b> · '
                f'고정비 <b style="color:{T["warn"]}">{CK(월평균고정)}</b> {단위K}'
                + (f' · {보고월}월 말 현금 {CK(말현금)} 이면 <b>매출이 없어도 고정비만으로 약 {버팀:,.1f}개월</b>'
                   if 버팀 is not None and 말현금 > 0 else '')
                + '</div>'
                f'<div class="scrollx"><table class="lined cfcomp cfsplit" style="width:{_너비:.3f}%; min-width:{자금최소폭 * _너비 / 100:.0f}px; table-layout:fixed;">{_콜}'
                f'<thead><tr><th>구분</th><th class="cum">실적 누적<br><span class="fcs">1~{보고월}월</span></th>'
                + ''.join(f'<th>{m}월</th>' for m in months)
                + f'<th>월평균</th></tr></thead><tbody>{_행}</tbody></table></div>'
                f'<details class="calcnote fold" style="margin-top:8px"><summary><b>어떻게 나눴나 · 쓰는 법</b> '
                f'<span class="muted">(누르면 펼쳐집니다)</span></summary><div class="assume" style="color:{T["ink2"]}">'
                f'<ul><li><b>변동비</b> = 매출에 따라 늘고 주는 돈 — 매입대금(매입채무 · 선급금 지급, 그리고 매입처에 낸 미지급금) · '
                f'운반비 · 판매수수료 · 외주용역(물류) · 수출제비용 · 광고 · 판촉</li>'
                f'<li><b>고정비</b> = 매출이 없어도 매달 나가는 돈 — 급여 · 4대보험 · 퇴직연금, 임차료 · 보험 · 통신 · 자문 · 지급수수료 · 이자 등 나머지</li>'
                f'<li>나눈 기준은 원장의 <b>상대 계정과목</b>과 거래처입니다 (가정 없음). 계정이 바뀌거나 새 거래처가 생기면 자동으로 다시 나뉩니다.</li>'
                f'<li><b>보는 법</b> : 고정비 월평균이 「매달 줄어드는 돈의 바닥」입니다. 입금 − 변동비(= 공헌 현금)가 고정비보다 커야 현금이 늘어납니다. '
                f'현금 ÷ 고정비 월평균 = 매출이 끊겨도 버틸 수 있는 개월 수(대략)입니다.</li></ul></div></details>')

        # ── 월말 잔액 꺾은선 : 아래 표의 월 칸 한가운데에 점이 오도록 %로 자리를 잡습니다
        #    (그림 영역도 표의 「1월~마지막 달」 칸과 정확히 같은 자리에 둡니다)
        #    맨 앞에는 기초 잔액을 얹어 「어디서 출발했는지」가 같이 보이게 합니다.
        #    기초 점은 표의 「구분」 머리칸 위쪽에 놓습니다.
        잔액K = ([기초현금 * 배율 / 표시나눔] + [v * 배율 / 표시나눔 for v in 기말]
                 + [v * 배율 / 표시나눔 for v in 예상기말])
        낮, 높 = min(잔액K), max(잔액K)
        폭 = (높 - 낮) or (abs(높) or 1.0)

        def 눈금만들기(lo, hi, 칸수=3):
            거친 = (hi - lo) / 칸수
            자리 = 10 ** math.floor(math.log10(거친)) if 거친 > 0 else 1
            for 배 in (1, 2, 2.5, 5, 10):
                if 거친 <= 자리 * 배:
                    간격 = 자리 * 배
                    break
            시작 = math.floor(lo / 간격) * 간격
            끝 = math.ceil(hi / 간격) * 간격
            값 = []
            v = 시작
            while v <= 끝 + 간격 * 1e-6:
                값.append(v)
                v += 간격
            return 값, 시작, 끝

        눈금값, 축낮, 축높 = 눈금만들기(min(0.0, 낮 - 폭 * 0.10), 높 + 폭 * 0.28)
        기초x = 자금이름비율 * 0.58        # 「구분」 머리칸 위쪽 — 기초 잔액이 서는 자리
        왼끝 = 기초x
        오른끝 = 자금이름비율 + 자금숫자비율 * (1 + len(months) + len(자금예상월))
        # 선 아래 옅은 면 — 계열색을 10%쯤으로 깔아 그림에 무게를 줍니다
        _주 = T['accent'].lstrip('#')
        _r, _g, _b = (int(_주[i:i + 2], 16) for i in (0, 2, 4))
        면색 = (f'linear-gradient(to bottom, rgba({_r},{_g},{_b},0.16) 0%,'
                f' rgba({_r},{_g},{_b},0.02) 100%)')

        def 가로(i):
            # i = 0 이면 기초, 그다음부터 1월 · 2월 … (표의 월 칸 한가운데)
            return 기초x if i == 0 else 자금이름비율 + 자금숫자비율 * (i + 0.5)     # 「실적 누적」 칸 다음부터

        def 세로(v):
            return (축높 - v) / (축높 - 축낮) * 100

        # 선(사다리꼴로 오려 낸 띠)과 그 아래 옅은 면
        굵기 = 1.2                                     # 선 두께의 절반 (px)
        바닥 = 세로(축낮)
        구간 = ''
        for i in range(len(잔액K) - 1):
            x1, x2 = 가로(i), 가로(i + 1)
            y1, y2 = 세로(잔액K[i]), 세로(잔액K[i + 1])
            자리 = f'left:{x1:.4f}%; width:{x2 - x1:.4f}%'
            구간 += (f'<div class="cfill" style="{자리}; background:{면색}; '
                     f'clip-path:polygon(0% {y1:.4f}%, 100% {y2:.4f}%, '
                     f'100% {바닥:.4f}%, 0% {바닥:.4f}%)"></div>')
            구간 += (f'<div class="cseg" style="{자리}; '
                     f'clip-path:polygon(0% calc({y1:.4f}% - {굵기}px), '
                     f'100% calc({y2:.4f}% - {굵기}px), '
                     f'100% calc({y2:.4f}% + {굵기}px), '
                     f'0% calc({y1:.4f}% + {굵기}px))"></div>')
        마지막 = len(잔액K) - 1
        표시 = ''.join(
            f'<div class="cdot{" last" if i == 마지막 else ""}'
            f'{" first" if i == 0 else ""}" '
            f'style="left:{가로(i):.4f}%; top:{세로(v):.4f}%"></div>'
            f'<div class="cval" style="left:{가로(i):.4f}%; top:{세로(v):.4f}%; '
            f'margin-top:-11px">{음수표기(f"{v:,.0f}")}</div>'
            for i, v in enumerate(잔액K))
        표시 += (f'<div class="ctag" style="left:{기초x:.4f}%; '
                 f'top:{세로(잔액K[0]):.4f}%">기초</div>')
        눈금 = ''.join(
            f'<div class="cgrid{" zero" if abs(t) < 1e-9 else ""}" '
            f'style="left:{왼끝:.4f}%; width:{오른끝 - 왼끝:.4f}%; top:{세로(t):.4f}%"></div>'
            f'<div class="cylab" style="width:calc({왼끝:.4f}% - 20px); '
            f'top:{세로(t):.4f}%">{t:,.0f}</div>'
            for t in 눈금값)
        축 = (f'<div class="cdomain" style="left:{왼끝:.4f}%"></div>'
              f'<div class="cbase" style="left:{왼끝:.4f}%; '
              f'width:{오른끝 - 왼끝:.4f}%; top:{세로(축낮):.4f}%"></div>')
        꺾은선 = (f'<div class="cashcap"><span class="t">월말 현금성 자산</span>'
                  f'<span class="u">기초 {CK(기초현금)} → {보고월}월 말 '
                  f'{CK(기말[-1])} {단위K}'
                  + (f' → {달표(자금예상월[-1])} 말 (예상) {CK(예상기말[-1])} {단위K}' if 예상기말 else '')
                  + '</span></div>'
                  f'<div class="cashline">'
                  f'{눈금}{축}{구간}{표시}</div>')

        예상글 = (f'{달표(자금예상월[0])} ~ {달표(자금예상월[-1])}' if 자금예상월 else '')
        예상안내 = (f' · <b>{예상글} (예상)</b> 은 앞으로 올 달입니다' if 자금예상월 else '')
        예상설명 = ('' if not 예상월들 else
                    f'<br><b>{예상월들[0]}~{예상월들[-1]}월(예상)</b> 은 실적이 아니라 '
                    f'<b>추정치</b>입니다 — 아래 붉은 칸의 가정 셋으로 계산했습니다.')

        # ── 예상 칸이 어떤 가정에서 나왔는지 (실적이 아니라 가정임을 붉게 적습니다)
        가정html = ''
        if 자금예상월:
            담당자넣음 = sorted({이름 for 이름, 칸 in 담당자칸.items()
                                 if any(m > 보고월 and v for m, v in 칸.items())})
            바뀐것 = []
            for k in 담당자넣음:
                첫달 = min(m for m, v in 담당자칸[k].items() if m > 보고월 and v)
                이름 = 자금이름.get(k, k)
                모형값 = {'AR': 가정['월매출'], 'AP': -가정['월원가'],
                          'PAY': -가정['월급여'], 'OTHOP': -가정['월기타']}.get(k)
                바뀐것.append(f'<b>{이름} {CK(담당자칸[k][첫달])}</b>'
                              + (f' <span class="muted">(가정값 {CK(모형값)} 대신)</span>'
                                 if 모형값 is not None else ''))
            담당자줄 = ('담당자가 「26년 자금」 3번 표에 넣으신 '
                        + ' · '.join(바뀐것)
                        + ' 은 가정보다 우선해 그대로 썼습니다.' if 담당자넣음 else
                        '설비 투자 · 유상증자 · 차입 · 상환은 회사가 정하는 일이라 넣지 않았습니다.')
            가정html = (
                f'<details class="calcnote fold" style="margin-top:12px"><summary><span style="color:{T["warn"]};font-weight:800">■ '
                f'{예상글} (예상)은 실적이 아니라 「가정」입니다</span> '
                f'<span class="muted">(누르면 가정 내용이 펼쳐집니다)</span></summary>'
                f'<div class="assume"><span class="red">■ '
                f'{예상글}은 실적이 아니라 「가정」입니다</span> — '
                f'가정은 셋뿐이고, 셋 다 <b>1~{보고월}월 실적을 그대로 나눈 값</b>이라 '
                f'새 원장을 올리면 자동으로 다시 계산됩니다 (단위 {단위K}).'
                f'<ul>'
                f'<li><b>매출액</b> 월 <span class="num">{CK(가정["월매출"])}</span> '
                f'= 1~{보고월}월 실적 {CK(sum(매출L))} ÷ {보고월}개월</li>'
                f'<li><b>매출원가</b> 월 <span class="num">{CK(가정["월원가"])}</span> '
                f'= 위 매출액 × 실적 원가율 {가정["원가율"] * 100:,.1f}%</li>'
                f'<li><b>판관비(현금)</b> 월 <span class="num">{CK(가정["월판관"])}</span> '
                f'= 실적 월평균 {CK(가정["월판관총"])}</li>'
                f'</ul>'
                f'→ 한 달 순증감 <span class="num">{CK(가정["월순증감"])}</span> '
                f'{단위K} = 매출액 − 매출원가 − 판관비(현금). '
                f'매출채권·매입채무·재고는 지금 수준 그대로라고 봐서 '
                f'<b>매출채권 회수 = 그 달 매출</b>, <b>매입채무 지급 = 그 달 매출원가</b> '
                f'로 놓았습니다.<br>{담당자줄}</div></details>')
        자금html = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card cfcard"><h3>현금흐름 <span class="unitbadge">단위 {단위K}</span></h3>
        <div class="sub">기초 현금에서 출발해 월별로 무엇 때문에
          늘고 줄었는지 나눴습니다 · 「실적 누적」 = 1~{보고월}월 실제 합계 (기초 · 기말은 1월 초 · {보고월}월 말 잔액){예상안내}</div>
        <style>
          .wrap .cfcard {{ container-type:inline-size; }}
          .wrap table.cfcomp th, .wrap table.cfcomp td {{ font-size:clamp(10.5px, 0.74cqw, 15px); padding-left:1px !important; padding-right:clamp(3px, 0.3cqw, 7px) !important;
              padding-top:clamp(5px, 0.42cqw, 9px) !important; padding-bottom:clamp(5px, 0.42cqw, 9px) !important; letter-spacing:-0.2px; }}
          .wrap table.cfcomp thead th {{ font-size:clamp(10.5px, 0.7cqw, 14px); }}
          .wrap table.cfcomp th .fcs {{ font-size:clamp(9.5px, 0.58cqw, 12px); }}
          .wrap .cfcard .cashline {{ height:clamp(126px, 12cqw, 240px); margin-bottom:clamp(12px, 1cqw, 20px); }}
          .wrap .cfcard .cashline .cval {{ font-size:clamp(10px, 0.7cqw, 14px); }}
          .wrap .cfcard .cashline .cylab, .wrap .cfcard .cashline .ctag {{ font-size:clamp(11px, 0.68cqw, 13.5px); }}
          .wrap .cfcard .cfsum {{ font-size:clamp(13px, 0.8cqw, 15px); }}
          .wrap table.cfcomp td.name {{ white-space:nowrap; padding-left:6px !important; }}
          .wrap table.cfcomp tr.sub td.name {{ padding-left:clamp(20px, 1.6cqw, 30px) !important; }}
          .wrap table.cfsplit tr.mid td, .wrap table.cfsplit tr.sub td.cum {{ font-weight:400; }}
          .wrap table.cfsplit tr.mid td {{ color:{T['ink']}; }}
          .wrap table.cfcomp td.cum {{ background:{T['navysoft']}; font-weight:700; }}
          .wrap table.cfcomp th.cum {{ background:{T['navy7']}; }}
          .wrap .cfsum {{ font-size:13px; color:{T['ink2']}; margin:4px 0 8px; }}
          .wrap table.cfsplit td.name {{ white-space:normal; }}
          .wrap table.cfcomp td:last-child {{ font-weight:400; color:{T['ink2']}; }}
          .wrap table.cfcomp tr.total td:last-child {{ font-weight:800; color:{T['ink']}; filter:none; }}
        </style>
        <div class="scrollx"><div style="min-width:{자금최소폭}px">
        {꺾은선}
        <table {자금표스타일}>{자금콜그룹}<thead><tr><th>구분</th>
        <th class="cum">실적 누적<br><span class="fcs">1~{보고월}월</span></th>
        {''.join(f'<th>{m}월</th>' for m in months)}
        {''.join(f'<th class="fc">{달이름(m)}<br><span class="fcs">{"예상" if m <= 12 else f"{(당해연도 + 1) % 100}년 예상"}</span></th>' for m in 자금예상월)}
        </tr></thead>
        <tbody>{자금rows}</tbody></table></div></div>
        {가정html}
        {고변html}</div></div>"""

    경고 = ''
    if False and 미분류건수:     # 실적보고 화면에서는 숨김 — 「데이터 점검」 에서만 표시 (2026-10-27)
        경고 = (f'<div class="warnbox">⚠ 판관비 {미분류건수:,}건이 활동분류 미입력 상태입니다. '
                f'"기타관리경비"로 집계되고 있습니다.</div>')

    상단html = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="note">모든 금액은 {'' if 통화 == 'KRW' else ('천 단위 ' if 표시나눔 == 1000 else '')}{단위K} 로 표시했습니다{' (천원 미만 반올림 · 월별 매출 추이 그래프만 백만원)' if 통화 == 'KRW' else ''} · 원장에서 직접 계산한 실제 수치입니다{환율문구}</div>
      {경고}
      {kpi}
      <div style="display:grid; grid-template-columns:1.5fr 1fr; gap:14px;">
        <div class="card"><h3>월별 매출 추이 <span class="unitbadge">단위 {'백만원' if 법인 == '씨티케이커넥트' else 단위K}</span></h3>
          <div class="sub">{(f'1~{보고월}월 · 맨 오른쪽 <b>1~{보고월}월 누적</b> {CK(누적26)} <span class="muted">(누적 막대 높이는 월별과 다른 기준)</span>') if 전년없음 else
            (f'1~{보고월}월 · {전년연도}년 동월 비교 &nbsp;|&nbsp; 맨 오른쪽 <b>1~{보고월}월 누적</b> = {전년연도}년 {CK(누적25)} ↔ {당해연도}년 {CK(누적26)}'
             '<span class="muted">(같은 기간끼리 견줍니다 · 누적 막대 높이는 월별과 다른 기준)</span>')}</div>
          {bars_line}
          <div class="legend">{'' if 전년없음 else f'<span><span class="sw" style="background:{SLATE};"></span> {전년연도}년</span>'}
          <span><span class="sw" style="background:{T['accent']};"></span> {당해연도}년</span></div></div>
        <div class="card"><h3>매출액 구성 <span class="unitbadge">단위 {단위K}</span></h3>
          <div class="sub">1~{보고월}월 누적 · 손익 계정과목 기준</div>{mix_html}</div>
      </div>
    </div>"""
    st.html(상단html)

    # ── 손익계산서 (왼쪽 3) + 항목별 거래처 세부내역 (오른쪽 1) ──────────
    # 선택상자는 카드 위 별도 줄에 둡니다 — 그래야 왼쪽·오른쪽 카드의 제목 줄이 나란히 맞습니다
    연도칸, 구분칸, 기준칸, 손익선택칸 = st.columns([1.25, 1.2, 1.45, 1.4], gap='small')
    with 연도칸:
        if 전년없음:      # CTKonnect : 25년 자료가 없어 연도 고르기를 숨깁니다
            st.session_state['dash_pl_year'] = f'{당해연도}년 (당해)'
        else:
            st.radio('연도', [f'{당해연도}년 (당해)', f'{전년연도}년 (전년)'], horizontal=True, key='dash_pl_year')
    with 구분칸:      # CTKonnect (2026-10-29) : 매출을 순액(장부) · 총액(거래액)으로 골라 봅니다
        st.radio('매출 기준', ['순액매출', '총액매출'], horizontal=True, key='dash_pl_gross',
                 disabled=not (당해보기 and 총액자료),
                 help='순액 = 장부 매출(중개 · IOR 은 수수료만) · 총액 = 중개 매입대금 · IOR 대납비까지 더한 거래액 '
                      '(연결패키지 요약보고 「4. 총액기준 월별 영업손익」 표)')
    with 기준칸:
        st.radio('판관비 구분', ['손익계산서 비용', '활동성 비용'], horizontal=True, key='dash_pl_kind',
                 help='판매관리비를 손익계산서 계정과목(상위 4개 + 기타)으로 볼지, 활동분류 항목으로 볼지 고릅니다')
        if 수불원가.get('직접재료비'):
            st.radio('손익 기준', [원가기준라벨, '퀵북 원장'], horizontal=True, key='dash_pl_basis',
                     disabled=not (당해보기 and not 활동기준),
                     help='원가 기준 = 재고수불부의 직접재료비 · 외주가공비 + 원장의 생산직 급여 · 제조 상각 '
                          '(당해 · 손익계산서 비용일 때만 고를 수 있습니다)')
    with 손익선택칸:
        손익골라본 = st.selectbox('계정과목', 드릴목록_손익, key='dash_pl_drill',
                                 help=f'고른 항목의 {연도X}년 1~{보고월X}월 누적 금액을 거래처별로 보여줍니다')
    손익조건, 손익부호 = 항목조건_손익(손익골라본)
    손익세부html = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>{손익골라본} <span class="unitbadge">단위 {단위K}</span></h3>
        <div class="sub">{연도X}년 1~{보고월X}월 누적 · 거래처별 · 상위 7곳 + 기타</div>
        {세부패널(누적IS[손익조건], '금액', 손익부호)}</div></div>"""
    손익예상안내 = ((' · <b>총액 기준</b> : 매출액 · 매출원가에 중개 매입대금 · IOR 대납비를 더한 거래액 (연결패키지 요약보고 4번 표)'
                     if 총액보기 else ' · <b>순액 기준</b> (장부 매출)')
                    + (f' · <b>{예상월들[0]}~{예상월들[-1]}월(예상)</b> 은 실적이 아니라 '
                    f'<b>가정</b>입니다' if 예상월들 else ''))
    # 예상 칸이 어떤 가정인지 — 현금흐름 표 아래 붉은 칸과 같은 가정임을 짚어 둡니다
    _예영익 = 예상영익원가 if 원가기준 else 예상영익
    손익가정 = ('' if not 예상월들 else
                f'<li class="red">■ 가정 : {예상월들[0]}~{예상월들[-1]}월 = 1~{보고월}월 월평균 '
                f'(월 매출 {CK(예상매출)} · 영업이익 {CK(_예영익)}) · 영업외손익 제외</li>')
    주석html = (f'<details class="calcnote fold"><summary><b>※ 주석</b> '
                f'<span class="muted">(누르면 펼쳐집니다)</span></summary><ul class="fnote">'
                f'{원가설명}{영업외설명}{손익가정}</ul></details>')

    # ── 개선과제 · 판관비 절감방안 (패키지 「개선과제」「판관비절감」 시트에 월별 칸으로 작성, 2026-10-27)
    #    당월 · 전월 · 연간(요약) 세 칸 · 누르면 접히고 펼쳐짐 · 지난 달 칸은 엑셀에서 잠김
    def 과제기록적용(시트키, 행들, 보고):
        """지난 달 칸은 「그 달이 당월이었을 때 올린 내용」 으로 고정합니다 (_보관자료/과제기록.json).
        · 당월 : 이번에 올린 패키지 내용으로 기록을 새로 씀
        · 지난 달 : 기록이 있으면 기록 그대로(패키지에서 고쳐도 반영 안 됨), 없으면 처음 한 번만 패키지 내용으로 기록
        · 지난 달을 고치려면 그 달 자료로 다시 올려(그 달이 당월) 기록을 바꾼 뒤 이번 달 자료를 올립니다"""
        자리 = os.path.join(보관폴더, '과제기록.json')
        try:
            with open(자리, encoding='utf-8') as f:
                전체 = json.load(f)
        except Exception:
            전체 = {}
        기록 = 전체.setdefault(시트키, {})
        키 = lambda r: f'{(r["연"] or 보고[0]):04d}-{r["월"]:02d}'
        당월 = f'{보고[0]:04d}-{보고[1]:02d}'
        묶음 = {}
        for r in 행들:
            묶음.setdefault(키(r), []).append(r)
        바뀜 = False
        if 기록.get(당월) != 묶음.get(당월, []):
            기록[당월] = 묶음.get(당월, [])
            바뀜 = True
        for k, rs in 묶음.items():
            if k < 당월 and k not in 기록:
                기록[k] = rs
                바뀜 = True
        if 바뀜:
            try:
                os.makedirs(보관폴더, exist_ok=True)
                with open(자리, 'w', encoding='utf-8') as f:
                    json.dump(전체, f, ensure_ascii=False, indent=1)
            except Exception:
                pass
        return [r for k in sorted(기록) if k <= 당월 for r in 기록[k]]

    def _과제줄(r):
        누락 = r.get('누락') or []
        쪽 = [f'<b>{_html.escape(r["구분"])}</b>' if r['구분'] else '']
        if r.get('업체'):
            쪽.append(f'{_html.escape(r["업체"])}' + (f' <span class="muted">({_html.escape(r["신규"])})</span>' if r.get('신규') else ''))
        본 = ' · '.join(x for x in 쪽 if x)
        뒤 = ' · '.join(x for x in [
            (f'기대 효과 월 {r["금액"]:,.0f} USD' if r['금액'] is not None else ''),
            (f'{_html.escape(r["시점"])}까지' if r.get('시점') else ''),
            _html.escape(r['상태']), _html.escape(r['담당'])] if x)
        return (f'<li style="margin:5px 0">{본}'
                + (f' — {_html.escape(r["항목"])}' if r['항목'] else '')
                + (f' — {_html.escape(r["내용"])}' if r.get('내용') else '')
                + (f' <span class="muted">({뒤})</span>' if 뒤 else '')
                + (f' <span style="color:{T["warn"]};font-weight:700">⚠ 누락 : {", ".join(누락)}</span>' if 누락 else '')
                + '</li>')

    def _과제키(r):
        return (r['연'] or 당해연도, r['월'])

    def _연간요약(행들, 당월글자수):
        """1월~당월 칸을 과제 구분별로 모아 짧게 — 길이는 당월 글 양에 맞추되 너무 짧으면 월 평균만큼."""
        올해 = [r for r in 행들 if _과제키(r)[0] == 당해연도 and _과제키(r) <= (당해연도, 보고월)]
        if not 올해:
            return ''
        # 같은 과제(과제 구분 + 업체명)는 가장 최근 달에 적은 것만 (달마다 다시 적어도 한 번만 셈)
        마지막달 = {}
        for r in 올해:
            k = (r['구분'], r.get('업체', '') or r['항목'][:12])
            마지막달[k] = max(마지막달.get(k, (0, 0)), _과제키(r))
        최근 = {i: r for i, r in enumerate(올해)
                if _과제키(r) == 마지막달[(r['구분'], r.get('업체', '') or r['항목'][:12])]}
        월글 = {}
        for r in 올해:
            월글[_과제키(r)] = 월글.get(_과제키(r), 0) + len(r['구분'] + r.get('업체', '') + r['항목'] + r.get('내용', ''))
        평균 = sum(월글.values()) / max(len(월글), 1)
        한도 = max(당월글자수, 평균, 80) * 1.2
        묶음 = {}
        for r in 최근.values():
            묶음.setdefault(r['구분'] or '기타', []).append(r)
        줄들, 쓴글 = [], 0
        for 구분 in sorted(묶음):
            rs = sorted(묶음[구분], key=lambda r: -(r['금액'] or 0))
            합 = sum(r['금액'] or 0 for r in rs)
            조각 = []
            for r in rs:
                t = (r.get('업체') or r['항목'][:14]) + (f'({r["신규"]})' if r.get('신규') else '')
                if r.get('시점'):
                    t += f' {r["시점"]}까지'
                if r['금액'] is not None:
                    t += f' {r["금액"] / 1000:,.0f}K'
                if 쓴글 + len(t) > 한도 and 조각:
                    조각.append(f'외 {len(rs) - len(조각)}건')
                    break
                조각.append(t); 쓴글 += len(t)
            줄들.append(f'<li style="margin:5px 0"><b>{_html.escape(구분)}</b> — {len(rs)}건'
                        + (f' · 기대 효과 계 월 {합:,.0f} USD' if 합 else '') + f'<br><span class="muted">{_html.escape(" · ".join(조각))}</span></li>')
        return ''.join(줄들)

    def _과제칸(제목, 키, 행들):
        줄들 = [r for r in 행들 if _과제키(r) == 키]
        목록 = ''.join(_과제줄(r) for r in 줄들)
        return (f'<div class="kpi-card"><div class="kpi-label">{제목} · {키[0]}년 {키[1]}월</div>'
                + (f'<ul style="margin:6px 0 0;padding-left:18px;font-size:13px;line-height:1.55">{목록}</ul>'
                   if 줄들 else '<div style="min-height:56px"></div>') + '</div>')

    def _과제섹션(제목, 머리말, 행들, 시트):
        당월키 = (당해연도, 보고월)
        전월키 = (당해연도, 보고월 - 1) if 보고월 > 1 else (당해연도 - 1, 12)
        당월글 = sum(len(r['구분'] + r.get('업체', '') + r['항목'] + r.get('내용', ''))
                     for r in 행들 if _과제키(r) == 당월키)
        연간 = _연간요약(행들, 당월글)
        지난 = sorted({_과제키(r) for r in 행들} - {당월키, 전월키}, reverse=True)
        지난html = ''.join(
            f'<div style="margin-top:8px"><b>{y}년 {m}월</b><ul style="margin:4px 0 0;padding-left:18px">'
            f'{"".join(_과제줄(r) for r in 행들 if _과제키(r) == (y, m))}</ul></div>' for y, m in 지난)
        return (f'{CARD_CSS}<div class="wrap" translate="no"><details class="fold" open>'
                f'<summary style="font-size:15px;font-weight:800;color:{T["ink"]};margin-bottom:6px">{제목} '
                f'<span class="muted" style="font-size:12.5px;font-weight:400">{머리말}</span></summary>'
                f'<div class="kpi-row" style="--n:3">'
                + _과제칸('당월', 당월키, 행들) + _과제칸('전월', 전월키, 행들)
                + (f'<div class="kpi-card"><div class="kpi-label">연간 · {당해연도}년 1~{보고월}월 요약</div>'
                   + (f'<ul style="margin:6px 0 0;padding-left:18px;font-size:13px;line-height:1.55">{연간}</ul>'
                      if 연간 else '<div style="min-height:56px"></div>') + '</div>')
                + '</div>'
                + (f'<details class="calcnote fold"><summary><b>지난 달 보기</b> '
                   f'<span class="muted">({len(지난)}개월 · 누르면 펼쳐집니다)</span></summary>{지난html}</details>'
                   if 지난 else '')
                + '</details></div>')

    def _기타섹션(행들):
        """기타주요사항 — 당월/전월 없이 지금 진행 중이거나 해야 할 일 목록. 완료는 접어 둠."""
        시트없음 = 행들 is None
        행들 = 행들 or []
        순서 = {'진행 중': 0, '해야 할 일': 1, '보류': 2}
        색 = {'진행 중': T['accent'], '해야 할 일': T['ink'], '보류': '#8a94a6', '완료': '#2e7d32'}
        def 줄(r):
            상태 = r['상태'] or '상태 없음'
            뒤 = ' · '.join(x for x in [(f'{_html.escape(r["기한"])}까지' if r['기한'] else ''), _html.escape(r['담당'])] if x)
            return (f'<li style="margin:6px 0"><span style="display:inline-block;min-width:64px;font-size:11.5px;font-weight:700;'
                    f'color:{색.get(상태, T["warn"])}">{_html.escape(상태)}</span>'
                    + (f'<b>{_html.escape(r["구분"])}</b> — ' if r['구분'] else '') + _html.escape(r['내용'])
                    + (f' <span class="muted">({뒤})</span>' if 뒤 else '')
                    + (f'<br><span class="muted" style="margin-left:64px">{_html.escape(r["비고"])}</span>' if r['비고'] else '')
                    + (f' <span style="color:{T["warn"]};font-weight:700">⚠ 누락 : {", ".join(r["누락"])}</span>' if r['누락'] else '')
                    + '</li>')
        할일 = sorted([r for r in 행들 if r['상태'] != '완료'], key=lambda r: 순서.get(r['상태'], 3))
        끝난 = [r for r in 행들 if r['상태'] == '완료']
        본문 = (f'<ul style="margin:6px 0 0;padding-left:18px;font-size:13px;line-height:1.55;list-style:none">{"".join(줄(r) for r in 할일)}</ul>'
                if 할일 else ('<div class="muted" style="min-height:40px;font-size:12.5px;margin-top:6px">올린 연결패키지에 '
                             '「기타주요사항」 시트가 없습니다 (v17 양식부터 있음)</div>' if 시트없음
                             else '<div style="min-height:40px"></div>'))
        return (f'{CARD_CSS}<div class="wrap" translate="no"><details class="fold" open>'
                f'<summary style="font-size:15px;font-weight:800;color:{T["ink"]};margin-bottom:6px">기타주요사항 '
                f'<span class="muted" style="font-size:12.5px;font-weight:400">조직운영안 외에 진행 중이거나 해야 할 일</span></summary>'
                f'<div class="kpi-card"><div class="kpi-label">진행 중 · 해야 할 일 {len(할일)}건</div>{본문}'
                + (f'<details class="calcnote fold"><summary><b>완료된 일</b> <span class="muted">({len(끝난)}건 · 누르면 펼쳐집니다)</span></summary>'
                   f'<ul style="margin:4px 0 0;padding-left:18px;list-style:none">{"".join(줄(r) for r in 끝난)}</ul></details>' if 끝난 else '')
                + '</div></details></div>')

    효율화html = ''
    if 당해보기 and 법인 == 'CTKonnect':      # 조직운영안 · 기타주요사항은 CTKonnect 패키지에만 있음
        _과제행 = 과제기록적용('조직운영안', (BS밑자료 or {}).get('원가절감') or [], (당해연도, 보고월))
        _최근 = {}
        for _r in _과제행:
            _최근[(_r['구분'], _r['항목'])] = _r
        _효과계 = sum(_r['금액'] or 0 for _r in _최근.values())
        효율화html = _과제섹션('조직운영안',
                              (f'과제 {len(_최근)}건 · 기대 효과 합계 월 {_효과계:,.0f} USD · 업무이관 · 조직 · 인력 · 프로세스 · 시스템'
                               if _최근 else f'{보고월}월 칸에 과제를 적어 올리시면 여기에 보입니다 (업무이관 · 조직 · 인력 · 프로세스 · 시스템)'),
                              _과제행, '조직운영안')
        효율화html += _기타섹션((BS밑자료 or {}).get('기타사항'))
    손익html = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>손익계산서 — {연도X}년 1~{보고월X}월 <span class="unitbadge">단위 {단위K}</span></h3>
        <div class="sub">{('비용은 양수 · <b>원가 기준(재고수불부)</b> : 매출액 → 변동비 → 공헌이익 → 고정비(제조간접비) → 매출이익 → 판관비(SR&amp;ED 차감 전 · 「판관비 합계」를 누르면 세부) → SR&amp;ED → 영업이익'
                           if 원가기준 else
                           ('비용은 양수 · 판관비는 <b>활동분류 항목</b>으로 나눠 표시했습니다' if 활동기준 else
                            '비용은 양수 · 판관비는 상위 4개 계정과 기타로 나눠 표시했습니다') + ' · 「판관비 합계」를 누르면 세부가 펼쳐집니다')} ·
          영업이익 아래 <b>영업외손익 · 세전이익</b>까지 보여줍니다{손익예상안내}</div>
        <input type="checkbox" class="togbox" id="pl-sga">
        <div class="scrollx"><table class="lined" {표스타일}>{콜그룹}<thead><tr><th>구분</th>
        {''.join(f'<th>{m}월</th>' for m in monthsX)}
        {''.join(f'<th class="fc">{m}월<br><span class="fcs">예상</span></th>' for m in 예상월들)}
        <th>누적</th></tr></thead>
        <tbody>{rows}</tbody></table></div>{주석html}</div></div>"""
    # ── 품목군별 손익 (약식) — 원가 기준 손익을 Resin · Straw · Cutlery 로 나눔 (2026-10-27 요청)
    품목군html = ''
    품목군자료 = (BS밑자료 or {}).get('품목군')
    if 원가기준 and 품목군자료:
        try:
            G = 품목군손익표(IS, monthsX, 매출L, C, 품목군자료)
        except Exception:
            G = None
        if G:
            def 군행(이름, 값들, 굵게=False):
                def f(v):
                    return '-' if abs(v * 배율 / 표시나눔) < 0.5 else CK(v)
                cls = ' class="total"' if 굵게 else ' class="sub"'
                return (f'<tr{cls}><td class="name" style="white-space:nowrap">{이름}</td>'
                        + ''.join(f'<td>{f(v)}</td>' for v in 값들)
                        + f'<td><b>{f(sum(값들))}</b></td></tr>')
            군줄 = ''
            총매출 = sum(sum(G[g]['매출']) for g in 품목군목록) or 1.0
            for k, g in enumerate(품목군목록, 1):
                o = G[g]
                군줄 += (f'<tr class="total hl"><td class="name" style="white-space:nowrap">{k}. {g} '
                         f'<span class="calc">(매출 비중 {sum(o["매출"]) / 총매출 * 100:.1f}%)</span></td>'
                         + '<td></td>' * (len(monthsX) + 1) + '</tr>')
                군줄 += 군행('매출액', o['매출'])
                군줄 += 군행('변동비', o['변동비'])
                군줄 += 군행('고정비', o['고정비'])
                군줄 += 군행('판관비', o['판관비'])
                군줄 += 군행('영업이익', o['영업이익'], True)
            칸수G = len(monthsX) + 1
            이름비G = 285 / (285 + 78 * 칸수G) * 100
            콜G = (f'<colgroup><col style="width:{이름비G:.4f}%">'
                   + f'<col style="width:{(100 - 이름비G) / 칸수G:.4f}%">' * 칸수G + '</colgroup>')
            품목군html = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>품목군별 손익 (약식) — {연도X}년 1~{보고월X}월 <span class="unitbadge">단위 {단위K}</span></h3>
        <div class="sub">위 손익계산서(원가 기준)를 품목군으로 나눈 것입니다 · 판관비는 SR&amp;ED 차감 전 ·
          세 품목군을 더하면 위 표와 같습니다 (영업이익은 SR&amp;ED 차감 전)</div>
        <div class="scrollx"><table class="lined" style="width:100%; max-width:{285 + 78 * 칸수G}px; table-layout:fixed;">
        {콜G}<thead><tr><th>구분</th>{''.join(f'<th>{m}월</th>' for m in monthsX)}<th>누적</th></tr></thead>
        <tbody>{군줄}</tbody></table></div>
        <div class="calcnote"><ul class="fnote">
          <li>매출액 = 원장 거래처별 매출 — 레진 판매 거래처(수불부 ④ 「레진판매」)는 Resin, 그 밖은 수불부 1~8월 매출 비율로 Straw · Cutlery</li>
          <li class="red">■ 가정 : 직접재료비 · 외주가공비는 수불부 품목군 원가율로, 직접노무비 · 고정비 · 판관비는 그 달 매출 비중으로 나눴습니다 (약식)</li>
        </ul></div></div></div>"""

    손익칸, 손익세부칸 = st.columns([2, 1], gap='small')   # 작은 화면에서도 오른쪽 세부내역이 읽히도록
    with 손익칸:
        st.html(손익html)
        if 효율화html:
            st.html(효율화html)
        if 품목군html:
            st.html(품목군html)
    with 손익세부칸:
        st.html(손익세부html)

    # ── 현금흐름 (왼쪽 3) + 항목별 거래처 세부내역 (오른쪽 1) ────────────
    if 자금표 is not None:
        누적자금상세 = 자금상세[자금상세['월'].le(보고월)] if 자금상세 is not None else None
        _, 자금선택칸 = st.columns([3, 1], gap='small')     # 카드 제목 줄을 맞추려고 선택상자만 따로
        with 자금선택칸:
            자금골라본 = st.selectbox('현금흐름 항목', [n for n, _, _ in 현금흐름목록], key='dash_cf_drill',
                                    help=f'고른 항목의 1~{보고월}월 누적 금액을 거래처별로 보여줍니다')
        _, 자금코드, 자금부호구분 = next(t for t in 현금흐름목록 if t[0] == 자금골라본)
        if 누적자금상세 is not None:
            자금대상 = 누적자금상세[누적자금상세['구분'].eq(자금코드)]
            if 자금부호구분 == '+':
                자금대상 = 자금대상[자금대상['영향'] > 0]
            elif 자금부호구분 == '-':
                자금대상 = 자금대상[자금대상['영향'] < 0]
        else:
            자금대상 = None
        자금세부html = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>{자금골라본} <span class="unitbadge">단위 {단위K}</span></h3>
        <div class="sub">1~{보고월}월 누적 · 거래처별 · 상위 7곳 + 기타</div>
        {세부패널(자금대상, '영향', 1)}</div></div>"""
        자금칸, 자금세부칸 = st.columns([3, 1], gap='small')
        with 자금칸:
            st.html(자금html)
            대출금옵션 = {}
            if 법인 == '씨티케이커넥트':          # 기초 = 26년 BS기초(25년 말 재무상태표)의 차입금 · 한도 없음
                try:
                    _b0 = pd.read_excel(io.BytesIO(데이터바이트), sheet_name='26년 BS기초', header=None)
                    _차입기초 = float(pd.to_numeric(_b0[_b0[5].astype(str).str.contains('차입금', na=False)][1], errors='coerce').fillna(0).sum())
                except Exception:
                    _차입기초 = 0.0
                _곳들 = (당해원장[당해원장['계정과목'].astype(str).str.contains('차입금')]['Name']
                         .dropna().astype(str).str.strip().unique().tolist())
                _곳 = _곳들[0] if len(_곳들) == 1 else '(25년 말 잔액 · 대출처 미상)'
                대출금옵션 = dict(기초보정=({_곳: _차입기초} if abs(_차입기초) > 0.5 else {}), 한도보정={},
                                  기초설명='기초 = 25년 말 재무상태표 차입금 (26년 BS기초 시트)', 통화='KRW')
            st.html(대출금html(당해원장, 보고월, 전년원장, **대출금옵션))
        with 자금세부칸:
            st.html(자금세부html)



# ══════════════════════════════════════════════════════════════
# 3-2. 월간실적보고 (보고서 5개 페이지)
# ══════════════════════════════════════════════════════════════
elif 메뉴 == '월간실적':
    페이지 = st.session_state['rep_page']
    월목록 = list(range(1, 보고월 + 1))
    단위 = 통화
    환율안내 = '' if 통화 == 'USD' else f' · 환율 1 USD = {환율:.4f} USD'

    # ── 공통 수치 ──────────────────────────────────────────
    매출N = [P.loc['매출액', m] for m in 월목록]
    원가N = [P.loc['매출원가', m] for m in 월목록]
    총이익N = [P.loc['매출총이익', m] for m in 월목록]
    판관N = [P.loc['판매관리비', m] for m in 월목록]
    SREDN = [float(SRED.get(m, 0.0)) for m in 월목록]           # 판관비에서 뺄 금액
    판관P = [판관N[i] - SREDN[i] for i in range(len(월목록))]     # 연결패키지 기준
    영익N = [총이익N[i] - 판관N[i] for i in range(len(월목록))]
    영익P = [총이익N[i] - 판관P[i] for i in range(len(월목록))]
    영외N = [P.loc['영업외손익', m] for m in 월목록]

    누계 = lambda a: float(np.nansum(a))
    R, CG, GP = 누계(매출N), 누계(원가N), 누계(총이익N)
    SG, SGP, SR = 누계(판관N), 누계(판관P), 누계(SREDN)
    OP, OPP, NO = 누계(영익N), 누계(영익P), 누계(영외N)
    PBT, PBTP = OP + NO, OPP + NO
    전매출 = float(전년동기.get('매출액', 0.0))
    전원가 = float(전년동기.get('매출원가', 0.0))
    전판관 = float(전년동기.get('판매관리비', 0.0))
    전총이익, 전영익 = 전매출 - 전원가, 전매출 - 전원가 - 전판관

    목표매출 = 입력값['월별'].get('매출 목표', [0.0] * 12)
    목표누계 = float(np.nansum(목표매출[:보고월]))

    def 카드(라벨, 값, 부가='', 델타='', 색='', 단위표기=None):
        """단위표기=None 이면 K USD/K USD, ''(빈칸)이면 단위 없이, 그 밖에는 그 글자로."""
        u = 단위K if 단위표기 is None else 단위표기
        단위html = f' <span class="unit">{u}</span>' if u else ''
        return (f'<div class="kpi-card"><div class="kpi-label">{라벨}</div>'
                f'<div class="kpi-value {색}" style="font-size:24px;font-weight:700">{값}{단위html}</div>'
                f'<div class="kpi-delta"><span class="muted">{부가}</span> {델타}</div></div>')

    def 표(제목, 부제, thead, rows, 높이=None):
        return (f'<div class="card"><h3>{제목}<span class="unitbadge">단위 {단위}</span></h3>'
                f'<div class="sub">{부제}</div>'
                f'<table class="lined"><thead><tr>{thead}</tr></thead>'
                f'<tbody>{rows}</tbody></table></div>')

    def 알림(제목, 본문):
        return (f'<div class="card"><h3>{제목}</h3><div class="sub" '
                f'style="margin:0;line-height:1.75;white-space:pre-line">{본문}</div></div>')

    def 서술(키, 기본=''):
        v = 입력값['서술'].get(키, '')
        return v if str(v).strip() else 기본

    st.write(f'### {페이지}')
    st.caption(f'{당해연도}년 {보고월}월 누적 · 원장(QuickBooks GL) 기준 · '
               f'표(금액) {단위} · 그래프 {단위K}{환율안내}')

    # ── 1. Executive Summary ───────────────────────────────
    if 페이지.startswith('1.'):
        달성 = f'목표 대비 {R / 목표누계 * 100:,.1f}%' if 목표누계 else '목표 미입력'
        html = f"""{CARD_CSS}<div class="wrap" translate="no">
          <div class="kpi-row" style="--n:4">
            {카드('매출액', CK(R), f'1~{보고월}월 누적',
                  증감HTML(R, 전매출, '전년동기') if 전매출 else '')}
            {카드('매출총이익', CK(GP), f'이익률 {GP / R * 100:,.1f}%' if R else '')}
            {카드('영업이익(손실)', CK(OPP), f'이익률 {OPP / R * 100:,.1f}%' if R else '')}
            {카드('현금 잔고', CK(현금잔고.get(보고월, 0.0)), f'{보고월}월 말 · 원장 계산',
                  증감HTML(현금잔고.get(보고월, 0.0), 현금잔고.get(보고월 - 1) if 보고월 > 1 else None))}
          </div>
          <div class="kpi-row" style="--n:4">
            {카드('매출원가', CK(CG), f'원가율 {CG / R * 100:,.1f}%' if R else '')}
            {카드('판관비 (SG&A)', CK(SG), f'매출 대비 {SG / R * 100:,.1f}%' if R else '')}
            {카드('영업외손익', CK(-NO), '이자 · 외환', '', 'warn' if NO < 0 else '')}
            {카드('세전이익', CK(PBTP), '연결패키지 기준')}
          </div>
          {알림('🔴 Red Flag — 리스크 · 이슈',
                서술('Red Flag', '실적파일의 「월간보고_입력」 시트에 적어주시면 여기에 나옵니다.'))}
          {알림('🟢 Green Flag — 성과 · 기회',
                서술('Green Flag', '실적파일의 「월간보고_입력」 시트에 적어주시면 여기에 나옵니다.'))}
          <div class="note">매출·이익·판관비·현금 잔고는 원장에서 자동 계산됩니다.
            Flag와 진척률처럼 원장에 없는 항목만 「월간보고_입력」 시트에서 읽습니다. · {달성}</div>
        </div>"""
        st.html(html)

        차트제목('월별 기말 현금 잔고')
        현금df1 = pd.DataFrame({'월': [f'{m}월' for m in 월목록],
                               '금액': [K(현금잔고.get(m, 0.0)) for m in 월목록]})
        기본c = alt.Chart(현금df1).encode(
            x=alt.X('월:N', sort=[f'{m}월' for m in 월목록], title=None,
                    axis=alt.Axis(labelAngle=0)),
            y=Y축(현금df1['금액'], 단위K),
            tooltip=[alt.Tooltip('월:N'), alt.Tooltip('금액:Q', format=',.0f', title=단위K)])
        st.altair_chart((기본c.mark_bar(cornerRadiusEnd=4, size=38, color=TEAL)
                         + 값라벨(기본c, dy=-9)).properties(height=320),
                        use_container_width=True)
        st.caption(f'원장의 현금·예금 계정에서 계산했습니다 · 기초 {CK(기초현금, 괄호=False)} '
                   f'{단위K} → {보고월}월 말 '
                   f'{CK(현금잔고.get(보고월, 0.0), 괄호=False)} {단위K}')

    # ── 2. P&L ─────────────────────────────────────────────
    elif 페이지.startswith('2.'):
        계획 = 입력값['월별']
        def 계획누계(k):
            v = 계획.get(k)
            if not v:
                return None
            t = float(np.nansum(v[:보고월]))
            return t if abs(t) > 0.005 else None

        행 = [('매출액', R, 계획누계('매출 목표'), 전매출, 0),
              ('매출원가', CG, 계획누계('매출원가 목표'), 전원가, 0),
              ('매출총이익', GP, None, 전총이익, 1),
              ('판관비 (SG&A)', SG, 계획누계('판관비 목표'), 전판관, 0),
              ('영업이익', OP, 계획누계('영업이익 목표'), 전영익, 1),
              ('영업외손익 (비용)', -NO, None, None, 0),
              ('세전이익', PBT, None, None, 1)]
        rows = ''
        for 이름, 실적, 계획v, 전년v, 굵게 in 행:
            차 = ('-' if 계획v in (None, 0) else C(실적 - 계획v))
            전 = '-' if 전년v is None else C(전년v)
            증 = (f'{(실적 / 전년v - 1) * 100:,.1f}%'
                 if 전년v and 전년v > 0 else '-')
            cls = ' style="font-weight:700"' if 굵게 else ''
            rows += (f'<tr{cls}><td>{이름}</td><td>{C(실적)}</td>'
                     f'<td>{"-" if 계획v is None else C(계획v)}</td><td>{차}</td>'
                     + ('' if 전년없음 else f'<td>{전}</td><td>{증}</td>') + '</tr>')
        기본remark = '「월간보고_입력」 시트의 P&L Remark 칸에 적어 주시면 여기에 나옵니다.'
        머리 = ('<th>구분</th><th>실적</th><th>계획</th><th>계획 대비</th>'
                + ('' if 전년없음 else f'<th>{전년연도}년 동기</th><th>전년 대비</th>'))
        html = f"""{CARD_CSS}<div class="wrap" translate="no">
          {표('P&L 요약', f'{당해연도}년 1~{보고월}월 누적 · 단위 {단위}', 머리, rows)}
          {알림('💡 P&L Remark', 서술('P&L Remark', 기본remark))}
        </div>"""
        st.html(html)

        긴 = pd.DataFrame({'월': [f'{m}월' for m in 월목록] * 2,
                           '구분': ['매출액'] * len(월목록) + ['영업이익'] * len(월목록),
                           '금액': [K(v) for v in 매출N + 영익P]})
        기본 = alt.Chart(긴).encode(
            x=alt.X('월:N', sort=[f'{m}월' for m in 월목록], title=None,
                    axis=alt.Axis(labelAngle=0)),
            xOffset=alt.XOffset('구분:N', sort=['매출액', '영업이익']),
            y=Y축(긴['금액'], 단위K),
            tooltip=[alt.Tooltip('월:N'), alt.Tooltip('구분:N'),
                     alt.Tooltip('금액:Q', format=',.0f', title=단위K)])
        막대 = 기본.mark_bar(cornerRadiusEnd=3, size=17).encode(
            color=alt.Color('구분:N', sort=['매출액', '영업이익'], scale=alt.Scale(
                domain=['매출액', '영업이익'], range=[BLUE, AMBER]), title=None))
        차트제목('월별 매출 · 영업이익')
        st.altair_chart((막대 + 값라벨(기본)).properties(height=330),
                        use_container_width=True)

    # ── 3. SG&A ────────────────────────────────────────────
    elif 페이지.startswith('3.'):
        구성 = pd.DataFrame({
            '항목': 판관비항목,
            '금액': [K(float(누적(P).get(it, 0.0))) for it in 판관비항목]})
        구성 = 구성[구성['금액'].abs() > 0.0005].sort_values('금액', ascending=False)
        c1, c2 = st.columns([1, 1])
        with c1:
            b1 = alt.Chart(구성).encode(
                y=alt.Y('항목:N', sort=list(구성['항목']), title=None),
                x=X축(구성['금액'], 단위K),
                tooltip=[alt.Tooltip('항목:N'), alt.Tooltip('금액:Q', format=',.0f', title=단위K)])
            차트제목('판관비 구성 — 활동분류(보고항목)')
            st.altair_chart((b1.mark_bar(color=TEAL, cornerRadiusEnd=4, size=16)
                             + 값라벨(b1, dy=0, dx=7, align='left')).properties(height=380),
                            use_container_width=True)
        with c2:
            계정 = (당해원장[당해원장['분류'].eq('판관비') & 당해원장['월'].le(보고월)]
                    .groupby('계정과목')['금액'].sum().mul(-배율 / 표시나눔)
                    .sort_values(ascending=False).head(10).reset_index())
            b2 = alt.Chart(계정).encode(
                y=alt.Y('계정과목:N', sort=list(계정['계정과목']), title=None),
                x=X축(계정['금액'], 단위K),
                tooltip=[alt.Tooltip('계정과목:N'), alt.Tooltip('금액:Q', format=',.0f', title=단위K)])
            차트제목('판관비 구성 — 계정과목 상위 10')
            st.altair_chart((b2.mark_bar(color=BLUE, cornerRadiusEnd=4, size=16)
                             + 값라벨(b2, dy=0, dx=7, align='left')).properties(height=380),
                            use_container_width=True)

        추이 = pd.DataFrame({
            '월': [f'{m}월' for m in range(1, 13)] * 2,
            '연도': [f'{당해연도}년'] * 12 + [f'{전년연도}년'] * 12,
            '금액': [(K(P.loc['판매관리비', m]) if m <= 보고월 else None) for m in range(1, 13)]
                    + [(None if 전년없음 else K(P전.loc['판매관리비', m])) for m in range(1, 13)]})
        기본2 = alt.Chart(추이.dropna()).encode(
            x=alt.X('월:N', sort=[f'{m}월' for m in range(1, 13)], title=None,
                    axis=alt.Axis(labelAngle=0)),
            y=Y축(추이['금액'], 단위K),
            color=alt.Color('연도:N', scale=alt.Scale(
                domain=[f'{당해연도}년', f'{전년연도}년'], range=[TEAL, SLATE]), title=None),
            tooltip=[alt.Tooltip('월:N'), alt.Tooltip('연도:N'),
                     alt.Tooltip('금액:Q', format=',.0f', title=단위K)])
        올해선 = 기본2.transform_filter(alt.datum.연도 == f'{당해연도}년')
        차트제목('월별 판관비 추이  (원장 기준)')
        st.altair_chart((기본2.mark_line(point=alt.OverlayMarkDef(size=70), strokeWidth=2.5)
                         + 값라벨(올해선, dy=-13)).properties(height=310),
                        use_container_width=True)
        st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
          {알림('💡 SG&A Remark', 서술('SG&A Remark',
                '전월·계획 대비 변동 원인을 「월간보고_입력」 시트에 적어주시면 여기에 나옵니다.'))}
        </div>""")

    # ── 4. 운영 KPI ────────────────────────────────────────
    elif 페이지.startswith('4.'):
        단일 = 입력값['단일']
        번율 = 단일.get('Burn Rate (월, USD)')
        if 번율 in (None, 0):
            번율 = -(현금잔고.get(보고월, 0.0) - 현금잔고.get(1, 0.0)) / max(보고월 - 1, 1)
            번율출처 = '2~%d월 평균 현금 감소 (자동 계산)' % 보고월
        else:
            번율출처 = '입력값'
        런웨이 = (현금잔고.get(보고월, 0.0) / 번율) if 번율 and 번율 > 0 else None

        def 값(k, 단위표시='%'):
            v = 단일.get(k)
            return f'{v:,.1f}{단위표시}' if isinstance(v, (int, float)) else '입력 필요'

        html = f"""{CARD_CSS}<div class="wrap" translate="no">
          <div class="kpi-row" style="--n:4">
            {카드('현금 잔고', CK(현금잔고.get(보고월, 0.0)), f'{보고월}월 말 · 원장 계산',
                  증감HTML(현금잔고.get(보고월, 0.0), 현금잔고.get(보고월 - 1) if 보고월 > 1 else None))}
            {카드('Burn Rate', CK(번율), 번율출처)}
            {카드('Runway', f'{런웨이:,.1f}' if 런웨이 else '-',
                  '현금 ÷ Burn Rate', 단위표기='개월')}
            {카드('판관비 중 인건비', f"{누적(P).get('인건비성 항목', 0.0) / SG * 100:,.1f}%" if SG else '-',
                  '원장 계산', 단위표기='')}
          </div>"""
        # CTKonnect : 캐나다의 R&D 집행률 · 파일럿 수율 · R&D 마일스톤 카드는 사업과 맞지 않아 뺐습니다
        진척 = 입력값['진척']
        if 진척:
            rows = ''.join(f'<tr><td>{k}</td><td>{v:,.0f}%</td></tr>' for k, v in 진척.items())
            html += 표('개발 · 운영 진척', '「월간보고_입력」 시트의 값입니다',
                       '<th>항목</th><th>진척률</th>', rows)
        else:
            html += 알림('개발 · 운영 진척',
                         '원장에 없는 값입니다. 실적파일 「월간보고_입력」 시트의 '
                         '「■ 진척률」 구역에 항목과 % 를 적어주시면 여기에 표로 나옵니다.')
        html += 알림('💡 CTKonnect Remark',
                     서술('KPI Remark', '운영 · 거래 진행 상황을 '
                          '「월간보고_입력」 시트에 적어주시면 여기에 나옵니다.'))
        html += '</div>'
        st.html(html)

        현금df = pd.DataFrame({'월': [f'{m}월' for m in 월목록],
                              '금액': [K(현금잔고.get(m, 0.0)) for m in 월목록]})
        기본3 = alt.Chart(현금df).encode(
            x=alt.X('월:N', sort=[f'{m}월' for m in 월목록], title=None,
                    axis=alt.Axis(labelAngle=0)),
            y=Y축(현금df['금액'], 단위K),
            tooltip=[alt.Tooltip('월:N'), alt.Tooltip('금액:Q', format=',.0f', title=단위K)])
        차트제목(f'월별 기말 현금 잔고  (기초 {CK(기초현금)} {단위K})')
        st.altair_chart((기본3.mark_bar(cornerRadiusEnd=4, size=34, color=TEAL)
                         + 값라벨(기본3, dy=-9)).properties(height=330),
                        use_container_width=True)

    # ── 5. Sales Forecast ──────────────────────────────────
    else:
        매출행 = 당해원장[당해원장['분류'].eq('매출') & 당해원장['월'].le(보고월)]
        거래처 = (매출행.groupby(매출행['Name'].fillna('(거래처 미기재)'))
                  .agg(누적=('금액', 'sum')).reset_index()
                  .rename(columns={'Name': '고객사'}).sort_values('누적', ascending=False))
        당월 = 매출행[매출행['월'].eq(보고월)].groupby(
            매출행['Name'].fillna('(거래처 미기재)'))['금액'].sum()
        전월s = 매출행[매출행['월'].eq(보고월 - 1)].groupby(
            매출행['Name'].fillna('(거래처 미기재)'))['금액'].sum() if 보고월 > 1 else pd.Series(dtype=float)
        고객목표 = 입력값['고객목표']

        rows = ''
        for i, r in enumerate(거래처.head(12).itertuples(), 1):
            이름 = getattr(r, '고객사')
            c, p = float(당월.get(이름, 0.0)), float(전월s.get(이름, 0.0))
            t = float(고객목표.get(이름, 0) or 0)
            증감 = '-' if abs(c - p) < 0.5 else (
                ('▲ ' if c - p > 0 else '▼ ') + C(abs(c - p)))
            색 = '' if abs(c - p) < 0.5 else ('up' if c - p > 0 else 'down')
            rows += (f'<tr><td>{i}</td>'
                     f'<td class="lft notranslate" translate="no">{이름}</td><td>{C(c)}</td>'
                     f'<td>{C(t) if t else "-"}</td>'
                     f'<td>{f"{c / t * 100:,.0f}%" if t else "-"}</td>'
                     f'<td class="{색}">{증감}</td>'
                     f'<td>{C(r.누적)}</td>'
                     f'<td>{r.누적 / R * 100:,.1f}%</td></tr>')
        rows += (f'<tr style="font-weight:700"><td></td><td class="lft">합계 (전체 거래처)</td>'
                 f'<td>{C(P.loc["매출액", 보고월])}</td><td>-</td><td>-</td>'
                 f'<td>-</td><td>{C(R)}</td><td>100.0%</td></tr>')
        html = f"""{CARD_CSS}<div class="wrap" translate="no">
          {표('주요 고객사별 실적', f'{보고월}월 · 원장의 거래처(Name) 기준 · 단위 {단위}',
              '<th>No</th><th class="lft">고객사</th><th>이번달</th><th>이번달 목표</th><th>달성률</th>'
              '<th>전월 대비</th><th>누적</th><th>비중</th>', rows)}
        </div>"""
        st.html(html)

        계열 = []
        for m in range(1, 13):
            if m <= 보고월:
                계열.append({'월': f'{m}월', '구분': f'{당해연도}년 실적',
                            '금액': K(P.loc['매출액', m])})
            if 목표매출 and 목표매출[m - 1]:
                계열.append({'월': f'{m}월', '구분': '목표', '금액': K(목표매출[m - 1])})
            if not 전년없음:
                계열.append({'월': f'{m}월', '구분': f'{전년연도}년',
                            '금액': K(P전.loc['매출액', m])})
        추이 = pd.DataFrame(계열)
        도메인 = [f'{당해연도}년 실적', f'{전년연도}년']
        색상 = [BLUE, SLATE]
        if (추이['구분'] == '목표').any():
            도메인 = [f'{당해연도}년 실적', '목표', f'{전년연도}년']
            색상 = [BLUE, AMBER, SLATE]
        if 전년없음:
            도메인, 색상 = 도메인[:-1], 색상[:-1]
        기본4 = alt.Chart(추이).encode(
            x=alt.X('월:N', sort=[f'{m}월' for m in range(1, 13)], title=None,
                    axis=alt.Axis(labelAngle=0)),
            xOffset=alt.XOffset('구분:N', sort=도메인),
            y=Y축(추이['금액'], 단위K),
            tooltip=[alt.Tooltip('월:N'), alt.Tooltip('구분:N'),
                     alt.Tooltip('금액:Q', format=',.0f', title=단위K)])
        실적만 = 기본4.transform_filter(alt.datum.구분 == f'{당해연도}년 실적')
        차트제목('월별 매출 — 실적 · 목표 · 전년')
        st.altair_chart((기본4.mark_bar(cornerRadiusEnd=3, size=13).encode(
            color=alt.Color('구분:N', sort=도메인, scale=alt.Scale(
                domain=도메인, range=색상), title=None))
            + 값라벨(실적만, dy=-8)).properties(height=330),
            use_container_width=True)
        if not 목표매출 or not any(목표매출):
            st.caption('목표(Plan)는 실적파일 「월간보고_입력」 시트의 '
                       '「■ 월별 목표」 구역에 넣으시면 자동으로 표시됩니다.')

# ══════════════════════════════════════════════════════════════
# 4. 미수채권 관리
# ══════════════════════════════════════════════════════════════
elif 메뉴 == '미수채권 관리':
    st.write(f'### 미수채권 관리 — {기준일.date()} 기준')
    st.caption('단위 USD · 만기 = 청구한 달의 다음 달 말일 · 거래처마다 오래된 청구부터 회수된 것으로 봅니다 (선입선출)')

    ar_all = 당해원장[당해원장['계정과목'].eq('매출채권')].copy()
    if ar_all.empty:
        st.warning('매출채권(Accounts Receivable) 계정을 찾지 못했습니다.')
        st.stop()
    기초채권 = float((BS밑자료 or {}).get('전기', {}).get('매출채권', 0.0) or 0.0)
    # 매출구분 — 같은 전표의 매출 계정으로 (중개 51103 · IOR 51104 · 상품 51101, 둘 이상이면 금액 큰 쪽)
    _매출 = 당해원장[당해원장['분류'].eq('매출')].copy()
    _매출['구'] = np.select([_매출['계정영문'].astype(str).str.startswith('51103'),
                            _매출['계정영문'].astype(str).str.startswith('51104'),
                            _매출['계정과목'].astype(str).isin(['상품매출', '제품매출'])],
                           ['상품수수료매출 (중개)', 'IOR매출', '상품매출'], '기타 매출')
    _매출['v'] = (_매출['Credit'] - _매출['Debit']).abs()
    구분맵 = (_매출.sort_values('v').groupby('TransactionID')['구'].last().to_dict())
    나이표, 초과회수 = 채권나이표(ar_all, 기초채권, 기준일, 구분맵)
    나이표['매출구분'] = 나이표['매출구분'].replace('', '기타 (대납비용 청구 · 이자 등)')
    if 나이표.empty:
        st.success('현재 미수 잔액이 있는 거래처가 없습니다.')
        st.stop()

    총미수 = float(나이표['채권총액'].sum())
    # CTKonnect : 1기 법인이라 기초채권이 없습니다 — 청구보다 더 받은 돈(초과입금)은 거래처별로 따로 적습니다
    _초과 = 나이표.attrs.get('초과거래처', {}) if 기초채권 < 0.5 else {}
    _장부 = (BS밑자료 or {}).get('당기', {}).get('매출채권')
    if 기초채권 >= 0.5:
        _초과글 = (f'<b>기초채권</b> {금액(기초채권)} 은 지난해에서 넘어온 몫이라 거래처를 붙이지 않았습니다. '
                   f'올해 청구보다 더 받은 돈 {금액(초과회수)} 은 이 기초채권을 갚은 것으로 보아 빼고 '
                   f'<b>{금액(max(기초채권 - 초과회수, 0))}</b> 만 남겼습니다.')
    else:
        _초과글 = ('1기 법인이라 기초채권은 없습니다. '
                   + ('청구보다 더 받은 돈(초과입금) — '
                      + ' · '.join(f'<b>{_html.escape(str(k))}</b> {금액(v)}' for k, v in _초과.items())
                      + ' — 은 미수에서 빼지 않고 따로 적었습니다 (다음 청구와 상계할 돈).' if _초과 else ''))
    _장부계산 = 총미수 - sum(_초과.values())
    if _장부 is not None and abs(float(_장부) - _장부계산) < 1:
        _초과글 += (f'<br>미수 합계 <b>{금액(총미수)}</b>' + (f' − 초과입금 {금액(sum(_초과.values()))}' if _초과 else '')
                    + f' = <b>{금액(_장부계산)}</b> 은 재무상태표의 {보고월}월말 매출채권과 같습니다.')
    elif _장부 is not None:
        _초과글 += (f'<br><span style="color:{T["warn"]}">재무상태표 {보고월}월말 매출채권 {금액(float(_장부))} 과 '
                    f'{금액(float(_장부) - _장부계산)} 차이가 납니다.</span>')
    정상 = float(나이표['within Due'].sum())
    연체 = 총미수 - 정상
    장기 = float(나이표[['~120', '121~']].sum().sum())

    kpi = f"""{CARD_CSS}<div class="wrap" translate="no"><div class="kpi-row" style="--n:3">
      <div class="kpi-card"><div class="kpi-label">{보고월}월 말 미수채권 합계</div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{금액(총미수)} <span class="unit">USD</span></div></div>
      <div class="kpi-card"><div class="kpi-label">기간 경과 (만기 지난 것)</div>
        <div class="kpi-value warn" style="font-size:24px;font-weight:700">{금액(연체)} <span class="unit">({연체/총미수*100:,.1f}%)</span></div>
        <div class="kpi-delta"><span class="muted">그중 91일 넘은 것 {금액(장기)}</span></div></div>
      <div class="kpi-card"><div class="kpi-label">만기 전</div>
        <div class="kpi-value ok" style="font-size:24px;font-weight:700">{금액(정상)} <span class="unit">({정상/총미수*100:,.1f}%)</span></div>
        <div class="kpi-delta"><span class="muted">미수 거래처 {나이표['거래처명'].nunique()}곳</span></div></div>
    </div></div>"""
    st.html(kpi)

    # ── 거래처별 채권 연령표 (금액이 큰 곳부터)
    구간합 = {c: float(나이표[c].sum()) for c in 채권구간}
    머리 = ''.join(f'<th>{c}</th>' for c in 채권구간)
    합계줄 = ''.join(f'<td>{금액(구간합[c])}</td>' for c in 채권구간)
    줄들 = ''
    기초색 = f' style="color:{T["ink3"]}"'
    _순서 = ['상품수수료매출 (중개)', 'IOR매출', '상품매출', '기타 매출', '기타 (대납비용 청구 · 이자 등)']
    for 구 in sorted(나이표['매출구분'].unique(), key=lambda g: _순서.index(g) if g in _순서 else 9):
        묶음 = 나이표[나이표['매출구분'] == 구].sort_values('채권총액', ascending=False)
        줄들 += (f'<tr class="total"><td class="lft">{_html.escape(구)}</td>'
                 + ''.join(f'<td>{금액(float(묶음[c].sum()))}</td>' for c in 채권구간)
                 + f'<td><b>{금액(float(묶음["채권총액"].sum()))}</b></td><td></td></tr>')
        for _i, 행 in 묶음.iterrows():
            칸 = ''.join(f'<td>{금액(float(행[c]))}</td>' for c in 채권구간)
            이름꼴 = 기초색 if str(행['거래처명']).startswith('기초채권') else ''
            줄들 += ('<tr class="sub"><td class="lft"' + 이름꼴 + '>&nbsp;&nbsp;&nbsp;'
                     + _html.escape(str(행['거래처명'])) + '</td>' + 칸
                     + '<td><b>' + 금액(float(행['채권총액'])) + '</b></td><td></td></tr>')
    st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>매출구분 · 거래처별 채권 연령표 <span class="unitbadge">단위 USD</span></h3>
        <div class="sub">매출구분 = 그 청구 전표의 매출 계정 (상품수수료매출(중개) · IOR매출 · 상품매출) · 구분 안에서는 채권총액이 큰 곳부터 · 만기 = 매출한 달의 다음 달 말일 ·
          경과일은 만기일에서 {기준일.date()} 까지</div>
        <div class="stick"><table class="lined" style="min-width:980px; margin-top:8px">
          <colgroup><col style="width:24%">{'<col style="width:9.5%">' * 6}
            <col style="width:11%"><col style="width:8%"></colgroup>
          <thead><tr><th>매출구분 / 거래처</th>{머리}<th>채권총액</th><th>대손충당금</th></tr></thead>
          <tbody>{줄들}
            <tr class="total"><td class="lft">합계</td>{합계줄}
              <td><b>{금액(총미수)}</b></td><td></td></tr>{''.join(
                f'<tr class="sub"><td class="lft" style="color:{T["ink3"]}">{_html.escape(str(k))} — 초과입금 (선수)</td>'
                + '<td></td>' * len(채권구간) + f'<td>{금액(-v)}</td><td></td></tr>' for k, v in _초과.items())}{(
                f'<tr class="total"><td class="lft">장부 매출채권</td>' + '<td></td>' * len(채권구간)
                + f'<td><b>{금액(_장부계산)}</b></td><td></td></tr>') if _초과 else ''}</tbody></table></div>
        <div class="calcnote"><b>어떻게 셌나</b> — 거래처마다 올해 청구(매출채권 차변)를 날짜순으로 쌓아 두고,
          들어온 돈(대변)으로 <b>오래된 것부터</b> 지웠습니다. 남은 청구만 만기일 기준으로 칸을 나눴습니다.<br>
          {_초과글}</div>
      </div></div>""")

    # ── 미수수익 (지연이자 등 아직 받지 않은 이자) — 매출채권과 따로 보여 줌 (2026-10-29)
    _미 = 당해원장[당해원장['계정과목'].astype(str).eq('미수수익')
                   | 당해원장['계정영문'].astype(str).str.contains('미수수익')].copy()
    if len(_미):
        _미['금액'] = _미['Debit'] - _미['Credit']
        _미['거래처'] = _미['Name'].map(lambda n: re.sub(r'\s*\((IC\s+)?(Customer|Customers|Vendor)\)\s*$', '', str(n)).strip()
                                        if n is not None and not (isinstance(n, float) and pd.isna(n)) else '(거래처 표시 없음)')
        _잔 = float(_미['금액'].sum())
        _줄 = ''.join(
            f'<tr class="sub"><td>{r["거래일"]:%Y-%m-%d}</td><td class="lft">{_html.escape(r["거래처"])}</td>'
            f'<td>{금액(r["금액"])}</td><td class="lft" style="font-weight:400;background:transparent">{_html.escape(str(r["Description"] or ""))}</td></tr>'
            for _, r in _미.sort_values('거래일').iterrows())
        _거 = _미.groupby('거래처')['금액'].sum()
        _거글 = ' · '.join(f'{_html.escape(k)} <b>{금액(v)}</b>' for k, v in _거.items() if abs(v) > 0.5)
        st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card" style="margin-top:14px"><h3>미수수익 — 지연이자 등 아직 받지 않은 이자 <span class="unitbadge">단위 USD</span></h3>
        <div class="sub">원장 미수수익 계정 · {보고월}월 말 잔액 <b>{금액(_잔)}</b> · 거래처별 : {_거글 or '-'} · 매출채권 연령표에는 들어가지 않습니다</div>
        <div class="scrollx"><table class="lined" style="width:100%; table-layout:fixed;">
          <colgroup><col style="width:12%"><col style="width:18%"><col style="width:14%"><col style="width:56%"></colgroup>
          <thead><tr><th>일자</th><th>거래처</th><th>금액</th><th>내용</th></tr></thead>
          <tbody>{_줄}<tr class="total"><td class="lft" colspan="2">합계 (잔액)</td><td>{금액(_잔)}</td><td></td></tr></tbody></table></div>
      </div></div>""")

    # ── 상품매출 건별 Aging (상품 판매분만, 2026-10-29)
    #    상품매출(51101) 이 함께 잡힌 전표의 매출채권 차변 = 상품 청구 1건.
    #    회수는 그 거래처의 매출채권 대변을 오래된 청구부터 지움 (선입선출 · 위 연령표와 같은 방식)
    _상품 = 상품채권html(당해원장, 기준일)
    if _상품:
        st.html(_상품.replace('<div class="card" style="margin-top:14px">',
                              '<details class="calcnote fold" style="margin-top:14px"><summary><b>상품매출 미수채권 건별 Aging</b> '
                              '<span class="muted">(상품 판매분 · 누르면 펼쳐집니다)</span></summary><div class="card">', 1)
                .replace('</table></div>\n      </div></div>', '</table></div>\n      </div></details></div>', 1))


# ══════════════════════════════════════════════════════════════
# 5. 당월 실적집계
# ══════════════════════════════════════════════════════════════
elif 메뉴 == '당월 실적보고':
    # ── 볼 달 고르기 : ◀ ▶ 로 옮기거나 숫자를 바로 골라도 됩니다 ─────
    if 'mo_month' not in st.session_state:
        st.session_state['mo_month'] = 보고월
    st.session_state['mo_month'] = min(max(int(st.session_state['mo_month']), 1), 보고월)

    def _달이동(걸음):
        st.session_state['mo_month'] = min(max(st.session_state['mo_month'] + 걸음, 1), 보고월)

    def _이번달로():
        st.session_state['mo_month'] = 보고월

    선택월 = st.session_state['mo_month']
    st.write(f'### {당해연도}년 {선택월}월 실적보고')
    st.caption(f'단위: USD · 비용 항목은 양수로 표시 · 원장 직접 집계 '
               f'· 원장에 {당해연도}년 1~{보고월}월 자료가 있습니다')

    앞칸, 달칸, 뒤칸, 오늘칸, 남은칸 = st.columns([0.5, 1.1, 0.5, 0.9, 5], gap='small')
    with 앞칸:
        st.button('◀', key='mo_prev', width='stretch', disabled=선택월 <= 1,
                  on_click=_달이동, args=(-1,), help='한 달 앞으로')
    with 달칸:
        st.selectbox('보는 달', list(range(1, 보고월 + 1)), key='mo_month',
                     format_func=lambda m: f'{m}월', label_visibility='collapsed')
    with 뒤칸:
        st.button('▶', key='mo_next', width='stretch', disabled=선택월 >= 보고월,
                  on_click=_달이동, args=(1,), help='한 달 뒤로')
    with 오늘칸:
        st.button('최종월', key='mo_last', width='stretch', disabled=선택월 == 보고월,
                  on_click=_이번달로, help=f'가장 마지막 달({보고월}월)로')
    선택월 = st.session_state['mo_month']

    이전월 = 선택월 - 1 if 선택월 > 1 else None
    기준선택 = st.radio('판관비 구분', ['손익계산서 비용', '활동성 비용'],
                        horizontal=True, key='mo_basis',
                        help='판매관리비를 손익계산서 계정과목으로 볼지, 활동분류로 볼지 고릅니다')
    손익기준_당월 = 기준선택.startswith('손익')
    표기준, 표기준전 = ((P손익, P손익전) if 손익기준_당월 else (P, P전))
    수준표 = 표기준.attrs['수준']
    보일줄 = [i for i in 표기준.index if i not in 화면제외]
    P보기, P전보기 = 표기준.loc[보일줄], 표기준전.reindex(표기준.index).loc[보일줄]
    표 = pd.DataFrame({
        '구분': 보일줄,
        f'당월({선택월}월)': P보기[선택월].values,
        f'전월({이전월}월)' if 이전월 else '전월': P보기[이전월].values if 이전월 else np.nan,
        f'전년동월({str(전년연도)[2:]}년 {선택월}월)': P전보기[선택월].values,
        f'당해누적({str(당해연도)[2:]}년 1월~{선택월}월)': 누적(P보기, upto=선택월).values,
        f'전년동기({str(전년연도)[2:]}년 1월~{선택월}월)': 누적(P전보기, upto=선택월).values,
    })
    표['전월 증감'] = 표.iloc[:, 1] - 표.iloc[:, 2]
    표['전년동월 증감'] = 표.iloc[:, 1] - 표.iloc[:, 3]
    표['누적 증감'] = 표.iloc[:, 4] - 표.iloc[:, 5]

    # ── 세 덩어리로 묶어 보여줍니다 (셋트마다 왼쪽에 구분선) ──────
    셋트 = [('전월대비', [표.columns[1], 표.columns[2], '전월 증감']),
            ('전년 동월대비', [표.columns[3], '전년동월 증감']),
            (f'{당해연도}년 {선택월}월 (누적)', [표.columns[4], 표.columns[5], '누적 증감'])]
    열순서 = [c for _, 열들 in 셋트 for c in 열들]
    증감열 = {'전월 증감', '전년동월 증감', '누적 증감'}
    첫열 = {열들[0] for _, 열들 in 셋트}          # 셋트가 시작되는 열 → 구분선

    def 두줄(이름):
        """「전년동월(25년 6월)」 → 「전년동월」 + 줄바꿈 + 「(25년 6월)」"""
        if 이름 in 증감열:
            return '증감액'
        i = 이름.find('(')
        return f'{이름[:i]}<br>{이름[i:]}' if i > 0 else 이름

    def 증감칸(v):
        """증감은 색으로 구분합니다 — 늘면 붉게, 줄면 푸르게, 변화 없으면 흐리게."""
        if v is None or pd.isna(v) or abs(v) < 0.5:
            return 'd zero', '-'
        return ('d plus', f'▲ {금액(abs(v))}') if v > 0 else ('d minus', f'▼ {금액(abs(v))}')

    그룹머리 = '<th class="gh gh0 name" rowspan="2">구분</th>'
    for 제목, 열들 in 셋트:
        그룹머리 += f'<th class="gh s1" colspan="{len(열들)}">{제목}</th>'
    소머리 = ''
    for c in 열순서:
        cls = ('d' if c in 증감열 else '') + (' s1' if c in 첫열 else '')
        소머리 += f'<th class="{cls.strip()}">{두줄(c)}</th>'

    rows = ''
    def 줄분류(이름):
        lv = 수준표.get(이름, 1)
        return 'total' if lv == 0 else ('sub' if lv == 1 else 'sub sub2')

    for _, r in 표.iterrows():
        cls = 줄분류(r['구분'])
        vals = ''
        for c in 열순서:
            선 = ' s1' if c in 첫열 else ''
            if c in 증감열:
                k, t = 증감칸(r[c])
                vals += f'<td class="{k}{선}">{t}</td>'
            else:
                vals += f'<td class="{선.strip()}">{금액(r[c])}</td>'
        rows += f'<tr class="{cls}"><td class="name">{r["구분"]}</td>{vals}</tr>'

    매출증감 = 표.loc[표['구분'] == '매출액', '전월 증감'].iat[0]
    매출전월 = P.loc['매출액', 이전월] if 이전월 else 0
    원가율 = (P.loc['매출원가', 선택월] / P.loc['매출액', 선택월] * 100) if P.loc['매출액', 선택월] else 0

    판관 = P.loc[판관비항목, [선택월]].copy()
    판관['전월'] = P.loc[판관비항목, 이전월] if 이전월 else 0
    판관['증감'] = 판관[선택월] - 판관['전월']
    상위 = 판관.reindex(판관['증감'].abs().sort_values(ascending=False).index).head(4)
    panel = ''
    for 항목, r in 상위.iterrows():
        방향 = '증가' if r['증감'] >= 0 else '감소'
        색 = 'up' if r['증감'] >= 0 else 'down'
        panel += (f'<tr><td>{항목}</td><td>{금액(r[선택월])}</td><td>{금액(r["전월"])}</td>'
                  f'<td class="{색}">{금액(r["증감"])} ({방향})</td>'
                  f'<td style="text-align:left; color:#94A3B8;">사유 입력 필요 — 회계 담당자 확인</td></tr>')

    표1 = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>1. {당해연도}년 {선택월}월 실적현황</h3>
        <div class="sub">전월 · 전년동월 · 누적을 셋트로 묶어 비교합니다 · 증감액은 색으로 구분</div>
        <div class="scrollx"><table class="fixed">
          <thead><tr>{그룹머리}</tr><tr>{소머리}</tr></thead>
          <tbody>{rows}</tbody></table></div></div></div>"""

    # ══ 오른쪽 : 항목을 고르면 거래처별 누적 내역 ═══════════════
    누적행 = 당해원장[당해원장['월'].le(선택월)]

    def 항목조건(이름):
        """표의 한 줄이 원장의 어떤 거래를 모은 것인지 알려 줍니다. (조건, 부호)"""
        d = 누적행
        if 이름 == '매출액':
            return d['분류'].eq('매출'), 1
        if 이름 == '상품매출':
            return d['계정과목'].isin(['제품매출', '상품매출']), 1
        if 이름 == '용역매출 (중개 · IOR)':
            return d['계정과목'].isin(['기타매출', '용역수익', '용역매출']), 1
        if 이름 == '매출원가':
            return d['분류'].eq('매출원가'), -1
        if 이름 in ('판매관리비', '판관비 (SR&ED 반영)'):
            return d['분류'].eq('판관비'), -1
        리스 = d['계정영문'].astype(str).str.strip().str.lower().eq('depreciation - rou - lease')
        상각 = d['분류'].eq('판관비') & (d['보고항목'].eq('상각비')
                                        | d['계정과목'].isin(['감가상각비', '무형자산상각비']))
        if 이름 == '리스상각비':
            return 상각 & 리스, -1
        if 이름 == '감가상각비':          # 리스분을 뺀 감가상각비
            return 상각 & ~리스, -1
        if 손익기준_당월:
            return d['분류'].eq('판관비') & d['계정과목'].eq(이름), -1
        return d['분류'].eq('판관비') & d['보고항목'].eq(이름), -1

    # 하위 항목을 거느린 줄(매출액 · 판매관리비 · 감가상각비)은 빼고 최하위 줄만 고르게 합니다
    드릴목록 = 최하위줄(보일줄, 수준표, 드릴제외)
    _, 선택칸 = st.columns([2, 1], gap='small')   # 카드 제목 줄을 맞추려고 선택상자만 따로 한 줄
    with 선택칸:
        골라본항목 = st.selectbox('계정과목', 드릴목록, key='mo_drill',
                                 help=f'고른 항목의 1~{선택월}월 누적 금액을 거래처별로 보여줍니다')

    조건, 부호 = 항목조건(골라본항목)
    대상 = 누적행[조건]

    옆 = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>{골라본항목} <span class="unitbadge">단위 USD</span></h3>
        <div class="sub">1~{선택월}월 누적 · 거래처별 · 상위 7곳 + 기타</div>
        {세부패널(대상, '금액', 부호)}</div></div>"""

    왼, 오 = st.columns([2, 1], gap='small')   # 작은 화면에서도 오른쪽 세부내역이 읽히도록
    with 왼:
        st.html(표1)
    with 오:
        st.html(옆)

    html = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>2. 매출 및 원가 요약</h3><div class="sub">자동 계산</div>
        <table><tbody>
          <tr><td>전월 대비 매출액</td><td>{금액(매출증감)} USD</td>
              <td>{증감HTML(P.loc['매출액', 선택월], 매출전월)}</td></tr>
          <tr><td>당월 원가율</td><td>{원가율:,.1f}%</td><td class="muted">매출원가 ÷ 매출액</td></tr>
          <tr><td>당월 영업이익률</td>
              <td>{(P.loc['영업이익(손실)',선택월]/P.loc['매출액',선택월]*100 if P.loc['매출액',선택월] else 0):,.1f}%</td>
              <td class="muted">영업이익 ÷ 매출액</td></tr>
        </tbody></table></div>
      <div class="card"><h3>3. 전월 대비 판관비 증감 상위 4개 항목</h3>
        <div class="sub">증감 사유는 시스템에서 알 수 없어 회계 담당자 입력이 필요합니다</div>
        <table><thead><tr><th>항목</th><th>당월</th><th>전월</th><th>증감</th>
        <th style="text-align:left;">증감 사유</th></tr></thead><tbody>{panel}</tbody></table></div>
    </div>"""
    st.html(html)

# ══════════════════════════════════════════════════════════════
# 6. 월별 실적집계
# ══════════════════════════════════════════════════════════════
elif 메뉴 == '월별 실적보고':
    st.write('### 월별 실적집계')

    ㄱ, ㄴ = st.columns(2)
    with ㄱ:
        연도선택 = st.radio('연도', [f'{당해연도}년 (당해)', f'{전년연도}년 (전년)'],
                           horizontal=True, key='mm_year')
    with ㄴ:
        기준선택 = st.radio('판관비 구분', ['손익계산서 비용', '활동성 비용'],
                           horizontal=True, key='mm_basis',
                           help='판매관리비를 손익계산서 계정과목으로 볼지, 활동분류로 볼지 고릅니다')
    손익기준 = 기준선택.startswith('손익')
    당해표 = P손익 if 손익기준 else P
    전년표 = P손익전 if 손익기준 else P전
    표원본 = 당해표 if 연도선택.startswith(str(당해연도)) else 전년표
    수준표 = 표원본.attrs['수준']
    당해보기 = 연도선택.startswith(str(당해연도))
    최대월 = 보고월 if 당해보기 else int(전년원장['월'].max() or 12)

    ms = list(range(1, 최대월 + 1))
    rows = ''
    def 줄분류(이름):
        lv = 수준표.get(이름, 1)
        return 'total' if lv == 0 else ('sub' if lv == 1 else 'sub sub2')

    for 항목 in [i for i in 표원본.index if i not in 화면제외]:
        cls = 줄분류(항목)
        cells = ''.join('<td class="{}">{}</td>'.format(
            's1' if m == ms[0] else '', 금액(표원본.loc[항목, m])) for m in ms)
        rows += (f'<tr class="{cls}"><td class="name">{항목}</td>{cells}'
                 f'<td class="s1">{금액(표원본.loc[항목, ms].sum())}</td></tr>')
    head = ''.join('<th class="{}">{}월</th>'.format(
        's1' if m == ms[0] else '', m) for m in ms)
    html = f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="note">단위: USD · 비용 항목은 양수로 표시</div>
      <div class="card"><h3>{연도선택} 월별 손익</h3>
        <div class="sub">원장 직접 집계 · 열너비를 모두 같게 맞췄습니다</div>
      <div class="scrollx"><table class="fixed"><thead><tr><th class="name">구분</th>{head}
        <th class="s1">누적</th></tr></thead><tbody>{rows}</tbody></table></div></div>
    </div>"""
    st.html(html)

    차트제목('최근 12개월 롤링 매출액 · 영업이익 추이', 'USD')
    시리즈 = []
    for i in range(11, -1, -1):
        m = 보고월 - i
        y, mm = (당해연도, m) if m >= 1 else (전년연도, m + 12)
        t = P if y == 당해연도 else P전
        시리즈.append({'라벨': f'{str(y)[2:]}.{mm:02d}', '매출액': t.loc['매출액', mm],
                       '영업이익': t.loc['영업이익(손실)', mm]})
    롤링 = pd.DataFrame(시리즈)
    긴 = 롤링.melt('라벨', var_name='구분', value_name='금액')
    st.altair_chart(alt.Chart(긴).mark_line(point=True).encode(
        x=alt.X('라벨:N', sort=list(롤링['라벨']), title=None,
                axis=alt.Axis(labelAngle=0)),
        y=alt.Y('금액:Q', title=None),
        color=alt.Color('구분:N', title=None, scale=alt.Scale(
            domain=['매출액', '영업이익'], range=[TEAL, AMBER]))),
        width='stretch', key='rolling12')

# ══════════════════════════════════════════════════════════════
# 7. 계정과목 상세  (신규)
# ══════════════════════════════════════════════════════════════
elif 메뉴 == '계정과목 상세':
    st.write('### 계정과목 상세')
    st.caption('계정과목을 고르면 월별 추이와 개별 거래내역까지 볼 수 있습니다.')

    c1, c2, c3 = st.columns([1, 1.4, 1])
    with c1:
        구분선택 = st.selectbox('구분', ['손익(IS) 전체', '매출', '매출원가', '판관비', '영업외손익',
                                       '재무상태(BS) 전체'], index=1, key='ad_cat')   # 처음엔 매출
    대상 = 당해원장.copy()
    if 구분선택 == '손익(IS) 전체':
        대상 = 대상[대상['계정분류'].eq('IS')]
    elif 구분선택 == '재무상태(BS) 전체':
        대상 = 대상[대상['계정분류'].eq('BS')]
    else:
        대상 = 대상[대상['분류'].eq(구분선택)]

    계정목록 = sorted(대상['계정과목'].dropna().unique())
    if not 계정목록:
        st.warning('해당 구분에 데이터가 없습니다.')
        st.stop()
    with c2:
        계정 = st.selectbox('계정과목', 계정목록, key='ad_acct',
                            index=계정목록.index('제품매출') if '제품매출' in 계정목록 else 0)   # 처음엔 제품매출
    with c3:
        월범위 = st.slider('월 범위', 1, 보고월, (1, 보고월), key='ad_month') if 보고월 > 1 else (1, 1)

    sel = 대상[(대상['계정과목'] == 계정) &
              (대상['월'] >= 월범위[0]) & (대상['월'] <= 월범위[1])].copy()
    부호 = 1 if 구분선택 == '매출' or 계정 in 매출계정 else -1
    sel['표시금액'] = sel['금액'] * 부호

    ms = list(range(월범위[0], 월범위[1] + 1))
    월합 = sel.groupby('월')['표시금액'].sum().reindex(ms, fill_value=0)
    전년sel = 전년원장[전년원장['계정과목'] == 계정]
    전년월합 = (전년sel.groupby('월')['금액'].sum() * 부호).reindex(ms, fill_value=0)

    합계 = 월합.sum()
    평균 = 월합.mean()
    최대월 = int(월합.abs().idxmax()) if len(월합) else 0
    전년합 = 전년월합.sum()

    kpi = f"""{CARD_CSS}<div class="wrap" translate="no"><div class="kpi-row" style="--n:4">
      <div class="kpi-card"><div class="kpi-label">기간 합계</div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{금액(합계)} <span class="unit">USD</span></div>
        <div class="kpi-delta">{증감HTML(합계, 전년합, f'{전년연도}년 동기대비')}</div></div>
      <div class="kpi-card"><div class="kpi-label">월 평균</div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{금액(평균)} <span class="unit">USD</span></div></div>
      <div class="kpi-card"><div class="kpi-label">최대 발생월</div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{최대월}월</div>
        <div class="kpi-delta"><span class="muted">{금액(월합.get(최대월, 0))} USD</span></div></div>
      <div class="kpi-card"><div class="kpi-label">거래 건수</div>
        <div class="kpi-value" style="font-size:24px;font-weight:700">{len(sel):,}건</div></div>
    </div></div>"""
    st.html(kpi)

    추이 = pd.concat([
        pd.DataFrame({'월': [f'{m}월' for m in ms], '금액': 월합.values, '연도': f'{당해연도}년'}),
    ] + ([] if 전년없음 else [pd.DataFrame({'월': [f'{m}월' for m in ms], '금액': 전년월합.values, '연도': f'{전년연도}년'})]))
    def _상위목록(묶음, 제목, 설명, 없을때):
        """상위 7개 + 기타 + 합계 — 이름 · 막대 · 금액(K USD) · 비중을 한 줄씩 (이름이 잘리지 않게)."""
        합 = float(묶음.sum())
        if not len(묶음) or abs(합) < 0.5:
            return (f'{CARD_CSS}<div class="wrap" translate="no"><div class="card"><h3>{제목}</h3>'
                    f'<div class="sub">{없을때}</div></div></div>')
        위 = 묶음.head(7)
        나머지 = float(묶음.iloc[7:].sum()) if len(묶음) > 7 else 0.0
        줄들 = list(위.items()) + ([(f'기타 ({len(묶음) - 7}곳)', 나머지)] if len(묶음) > 7 else [])
        최대 = max(abs(v) for _, v in 줄들) or 1.0
        몸 = ''.join(
            f'<div class="mix-item"><div class="top"><span>{_html.escape(str(이름))}</span>'
            f'<span><b>{v / 1000:,.1f}</b> <span class="muted">K USD ({v / 합 * 100:,.1f}%)</span></span></div>'
            f'<div class="mix-track"><div class="mix-fill" style="width:{abs(v) / 최대 * 100:.0f}%;"></div></div></div>'
            for 이름, v in 줄들)
        return (f'{CARD_CSS}<div class="wrap" translate="no"><div class="card"><h3>{제목} '
                f'<span class="unitbadge">단위 K USD</span></h3><div class="sub">{설명}</div>{몸}'
                f'<div class="mix-total">합계<span>{합 / 1000:,.1f} <span class="muted">K USD</span></span></div>'
                f'</div></div>')

    # 왼쪽 : 월별 추이(꺾은선) · 오른쪽 : 거래처별 상위 7 + 기타 (활동세부별 구성은 뺌, 2026-10-27)
    g1, g2 = st.columns([1.7, 1], gap='small')
    with g1:
        차트제목(f'{계정} 월별 추이 — {당해연도}년' + ('' if 전년없음 else f' · {전년연도}년'), 'K USD')
        추이['K'] = 추이['금액'] / 1000
        추이['글'] = 추이['K'].map(lambda v: f'{v:,.0f}')
        _x = alt.X('월:N', sort=[f'{m}월' for m in ms], title=None, axis=alt.Axis(labelAngle=0, labelFontSize=13))
        _색 = alt.Color('연도:N', title=None, scale=alt.Scale(
            domain=[f'{당해연도}년', f'{전년연도}년'], range=[T['accent'], SLATE]),
            legend=alt.Legend(orient='top', labelFontSize=13))
        _기본 = alt.Chart(추이).encode(x=_x, y=alt.Y('K:Q', title=None), color=_색)
        추이['위'] = 추이['K'].ge(추이.groupby('월')['K'].transform('max'))   # 큰 값 숫자는 위, 작은 값은 아래
        _기본 = alt.Chart(추이).encode(x=_x, y=alt.Y('K:Q', title=None), color=_색)
        st.altair_chart((_기본.mark_line(strokeWidth=3, point=alt.OverlayMarkDef(size=70, filled=True))
                         + _기본.transform_filter('datum.위').mark_text(dy=-12, fontSize=12, fontWeight='bold').encode(text='글:N')
                         + _기본.transform_filter('!datum.위').mark_text(dy=16, fontSize=12, fontWeight='bold').encode(text='글:N'))
                        .properties(height=420), width='stretch', key='ad_trend')
    with g2:
        거래처 = sel.groupby(sel['Name'].fillna('(거래처 미기재)'))['표시금액'].sum()
        거래처 = 거래처[거래처.abs() > 0.5].sort_values(ascending=False)
        st.html(_상위목록(거래처, '거래처별 상위 7 + 기타',
                         f'{월범위[0]}~{월범위[1]}월 누적 · 금액 큰 순서', '거래처명이 기록되지 않은 계정입니다.'))

    st.write(f'**{계정} 거래내역 ({월범위[0]}~{월범위[1]}월, {len(sel):,}건)**')
    상세 = sel[['거래일', '계정영문', 'Description', 'Name', '표시금액',
               '활동분류(대분류)', '활동세부', '분류상태', 'TransactionID']].copy()
    상세['거래일'] = 상세['거래일'].dt.date
    상세 = 상세.rename(columns={'계정영문': '원장 계정(영문)', 'Description': '적요',
                              'Name': '거래처', '표시금액': '금액(USD)'})
    st.dataframe(상세.sort_values('거래일'), width='stretch', hide_index=True,
                 column_config={'금액(USD)': st.column_config.NumberColumn(format='%,.2f')})
    st.download_button('이 내역 CSV로 내려받기',
                       상세.to_csv(index=False).encode('utf-8-sig'),
                       file_name=f'{당해연도}_{계정}_{월범위[0]}-{월범위[1]}월.csv',
                       mime='text/csv', key='ad_dl')

# ══════════════════════════════════════════════════════════════
# 8. 데이터 점검  (신규)
# ══════════════════════════════════════════════════════════════
elif 메뉴 == '데이터 점검':
    st.write('### 데이터 점검')
    st.caption('숫자를 보고에 쓰기 전에 이 페이지에서 원장이 제대로 읽혔는지 확인하세요.')

    검증 = []
    합계 = 당해원장['금액'].sum()
    검증.append(('차변 = 대변 (전체 합계 0)',
                 '0.00' if abs(합계) < 0.005 else f'{합계:,.2f}', abs(합계) < 1))
    미매핑 = int(당해원장['계정과목'].isna().sum())
    검증.append(('계정 매핑 누락 건수', f'{미매핑:,}건', 미매핑 == 0))
    범위밖 = int((~당해원장['월'].between(1, 12)).sum())
    검증.append(('월 범위(1~12) 이탈 건수', f'{범위밖:,}건', 범위밖 == 0))
    미래 = int((당해원장['월'] > 보고월).sum())
    검증.append((f'보고월({보고월}월) 이후 거래 건수 — 날짜 오류 신호', f'{미래:,}건', 미래 == 0))
    검증.append(('판관비 활동분류 미입력 건수', f'{미분류건수:,}건', 미분류건수 == 0))

    줄들 = ''.join(
        f'<div class="chk {"ok" if ok else "wn"}">'
        f'<span class="ic">{"✓" if ok else "!"}</span>'
        f'<span class="t">{이름}</span><span class="v">{값}</span></div>'
        for 이름, 값, ok in 검증)
    통과 = sum(1 for *_, ok in 검증 if ok)
    st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>점검 결과
        <span class="unitbadge">{통과}/{len(검증)}개 항목 정상</span></h3>
        <div class="sub">색을 못 알아보셔도 되도록 아이콘(✓ · !)을 함께 넣었습니다</div>
        {줄들}</div></div>""")

    st.write('---')
    st.write('**계정 매핑이 안 된 원장 계정**')
    미매핑표 = (당해원장[당해원장['계정과목'].isna()]
               .groupby('계정영문')['금액'].agg(건수='size', 합계='sum').reset_index())
    if len(미매핑표):
        st.dataframe(미매핑표, width='stretch', hide_index=True)
        st.warning('위 계정은 BS_IS_매핑 시트에 추가해야 손익에 반영됩니다. 회계 담당자 확인이 필요합니다.')
    else:
        st.success('모든 원장 계정이 매핑되어 있습니다.')

    분포 = 당해원장.groupby('월').size().reindex(range(1, 13), fill_value=0)
    칸 = ''.join(f'<td style="{"color:" + T["warn"] + ";font-weight:800" if (m > 보고월 and n) else ""}">'
                 f'{n:,}</td>' for m, n in 분포.items())
    st.html(f"""{CARD_CSS}<div class="wrap" translate="no"><div class="card"><h3>월별 거래 건수
      <span class="unitbadge">단위 건</span></h3>
      <div class="sub">보고월({보고월}월) 이후 달에 숫자가 있으면 날짜(일/월) 읽기 오류를 의심해야 합니다</div>
      <table style="width:100%;table-layout:fixed"><thead><tr>{''.join(f'<th>{m}월</th>' for m in range(1, 13))}<th>합계</th></tr></thead>
      <tbody><tr>{칸}<td><b>{int(분포.sum()):,}</b></td></tr></tbody></table></div></div>""")

    st.write('**분류 상태 요약**')
    상태 = (당해원장[당해원장['계정분류'].eq('IS')]['분류상태']
            .fillna('(표시 없음)').value_counts().reset_index())
    상태.columns = ['분류상태', '건수']
    st.dataframe(상태, width='stretch', hide_index=True)


# ══════════════════════════════════════════════════════════════
# 9. 연결재무제표 패키지  (틀만 만들어 둔 자리)
# ══════════════════════════════════════════════════════════════
elif 메뉴 in 패키지하위:
    # ── 연결재무제표 패키지 : 재무제표(BS · IS · MC) · 주석사항 (지금은 자리만 잡아 둡니다)
    설명 = {
        'BS':   ('BS  재무상태표', '연결 기준 자산·부채·자본',
                 [('원장 계정 → 패키지 계정 매핑', 'BS_IS_매핑 시트의 「계정분류 BS」'),
                  ('기초잔액 + 당기 증감 → 기말잔액', '26년 원장_raw · 26년 BS기초'),
                  ('연결 조정분개 반영', '대표님이 주실 연결 분개장'),
                  ('전분기 대비 증감 확인', '직전 분기 패키지')]),
        'IS':   ('IS  손익계산서', '연결 기준 수익·비용',
                 [('원장 계정 → 패키지 계정 매핑', 'BS_IS_매핑 시트의 「계정분류 IS」'),
                  ('매출 · 매출원가 · 판관비 집계', '월별 실적집계와 같은 기준'),
                  ('내부거래 제거 · 연결 조정', '대표님이 주실 연결 분개장'),
                  ('전년 동기 대비 확인', '25년 원장_raw')]),
        'MC':   ('MC  제조원가명세서', '재료비 · 노무비 · 경비 → 당기제품제조원가',
                 [('원장 계정 → 원가 요소 매핑', '재료비 · 노무비 · 제조경비'),
                  ('재공품 기초 · 기말 반영', '26년 BS기초 · 26년 원장_raw'),
                  ('당기제품제조원가 산출', 'IS 화면의 매출원가와 연결'),
                  ('연결 조정분개 반영', '대표님이 주실 연결 분개장')]),
        '주석사항': ('주석사항', '패키지에 붙는 주석',
                 [('현금및현금성자산 구성', '26년 자금 시트'),
                  ('매출채권 · 대손 현황', '미수채권 관리 화면'),
                  ('차입금 · 리스부채 명세', '26년 원장_raw'),
                  ('특수관계자 거래', '대표님 확인 필요')]),
    }
    if 메뉴 == '재무제표':
        머리 = '재무제표'
        st.html(f"""<div class="wrap" translate="no" style="margin-bottom:0">
  <div class="page-head"><span class="t">{머리}</span>
    <span class="pill">연결재무제표 패키지</span></div>
</div>""")
        왼칸, 가운데칸, 오른칸 = st.columns([1.5, 1.5, 1], gap='medium')
        with 왼칸:
            # 첫 칸(BS)이 처음부터 체크되어 있습니다
            고른탭 = st.radio('재무제표 구분', 재무제표탭, horizontal=True, key='pkg_tab',
                              help='보고 싶은 표를 고르세요 '
                                   '(BS 재무상태표 · IS 손익계산서 · MC 제조원가명세서)')
            구분 = 고른탭.split()[0]
            말 = st.radio('표기', 표기구분, horizontal=True, key='pkg_lang',
                          help='계정과목을 국문으로 볼지 영문으로 볼지 고르세요 '
                               '(연결패키지 BS·IS 시트의 국문·영문 열과 같습니다)')
            국문으로 = 말.startswith('국문')
        with 가운데칸:
            연결패키지양식칸()
        with 오른칸:
            연결패키지칸()
    else:
        국문으로 = True
        구분 = 메뉴
        st.html(f"""<div class="wrap" translate="no" style="margin-bottom:0">
  <div class="page-head"><span class="t">{메뉴}</span>
    <span class="pill">연결재무제표 패키지</span></div>
</div>""")
    이름, 한줄, 할일 = 설명[구분]
    st.caption(f'{한줄} · {당해연도}년 {(보고월 - 1) // 3 + 1}분기 기준 (원장 1~{보고월}월) · '
               + ('기준 자료인 연결패키지의 BS·IS 시트 숫자(원장 수식 · 본사 기초)를 그대로 보여 줍니다'
                  if (BS밑자료 or {}).get('패키지시트') else '아직 안쪽 내용은 비어 있습니다'))

    # ── 재무제표 표 그리기 ─────────────────────────────────────
    #    국문 : 연결패키지 BS·IS 서식 그대로
    #    영문 : 퀵북 계정 이름 그대로 (전기 · 당기 · 증감)
    그렸다 = False
    if 구분 in ('BS', 'IS') and BS밑자료:
        환칸1, 환칸2, _여백 = st.columns([1, 1, 2], gap='small')
        _환 = (BS밑자료 or {}).get('환율') or {}
        전기환율 = 환칸1.number_input(f'{당해연도 - 1}년 말 환율 (원/USD)', value=float(_환.get('전기') or 0.0),
                                      step=0.01, format='%.2f', key='bs_fx0')
        당기환율 = 환칸2.number_input(f'{당해연도}년 {보고월}월말 환율 (원/USD)', value=float(_환.get('당기') or 0.0),
                                      step=0.01, format='%.2f', key='bs_fx1')

        def _돈(v, 자리=2):
            if abs(v) < 0.005:
                return '0.00' if 자리 else '0'
            글 = f'{abs(v):,.{자리}f}'
            return f'<span class="neg">({글})</span>' if v < 0 else 글

        # ── 국문 서식 값 — 영문 화면에서도 「국문과 맞는지」 대사에 씁니다 ──
        서식 = BS서식 if 구분 == 'BS' else IS서식
        if 구분 == 'BS':
            전기값 = dict(BS밑자료['전기'])
            당기값 = dict(BS밑자료['당기'])
        else:
            전기값 = dict(BS밑자료['IS']['전기'])
            당기값 = dict(BS밑자료['IS']['당기'])
        푼값 = [None] * len(서식)

        def _금액(i, 어느):
            if 푼값[i] is not None and 어느 in 푼값[i]:
                return 푼값[i][어느]
            _번, 국, _영, _수준, 더할 = 서식[i]
            if 더할:
                v = sum(g * _금액(j, 어느) for j, g in 더할)
            else:
                v = float((전기값 if 어느 == '전기' else 당기값).get(국, 0.0))
            푼값[i] = (푼값[i] or {})
            푼값[i][어느] = v
            return v

        맞춘값 = {}
        if 구분 == 'BS':
            # 미처분이익잉여금은 원장에 없는 「맞추는 숫자」라 양쪽이 맞도록 채웁니다
            자산i = next((i for i, r in enumerate(서식) if r[1] == '자산총계'), None)
            합계i = next((i for i, r in enumerate(서식) if r[1] == '부채및자본총계'), None)
            잉여i = next((i for i, r in enumerate(서식) if r[1] == '미처분이익잉여금'), None)
            if None not in (자산i, 합계i, 잉여i):
                for 어느, 값들 in (('전기', 전기값), ('당기', 당기값)):
                    값들['미처분이익잉여금'] = 0.0
                    푼값[:] = [None] * len(서식)
                    차이 = _금액(자산i, 어느) - _금액(합계i, 어느)
                    값들['미처분이익잉여금'] = round(차이, 2)
                    맞춘값[어느] = round(차이, 2)
                푼값[:] = [None] * len(서식)

        def _서식줄(이름):
            i = next((k for k, r in enumerate(서식) if r[1] == 이름), None)
            return None if i is None else _금액(i, '당기')

        국문합 = ({'자산총계': _서식줄('자산총계'),
                   '부채및자본총계': _서식줄('부채및자본총계')} if 구분 == 'BS'
                  else {'매출액': _서식줄('매 출 액'), '매출원가': _서식줄('매 출 원 가'),
                        '당기순이익': _서식줄('당기순이익(손실)')})

        # ── 퀵북 표 (실적보고 엑셀작성에서 올려 둔 원장) ────────
        퀵 = None
        원장있음, 원장이름, 원장도장 = 보관원장정보()
        if 원장있음:
            try:
                퀵 = 퀵북표읽기(원장도장, 구분)
            except Exception:
                퀵 = None
        영문합 = {}
        if 퀵:
            찾기표 = dict(BS밑자료['영문BS'].get('전기맞춤', {})) if 구분 == 'BS' else {
                _맞춤이름(n): float(v)
                for n, v in BS밑자료['IS']['전기계정'].items()}
            줄들 = 퀵['줄']
            전기칸 = [None] * len(줄들)
            for i, x in enumerate(줄들):
                if x['당기'] is None:       # 금액이 없는 구역 머리글은 전기도 비워 둡니다
                    continue
                전기칸[i] = 찾기표.get(_맞춤이름(x['이름']))
            for i, x in enumerate(줄들):    # 표에 없는 Total 줄은 안쪽 줄을 더해 채웁니다
                if not x['총계'] or 전기칸[i] is not None:
                    continue
                깊 = x['깊이']
                모아, 봤다 = 0.0, False
                j = i - 1
                while j >= 0 and 줄들[j]['깊이'] > 깊:
                    if 줄들[j]['깊이'] == 깊 + 1 and 전기칸[j] is not None:
                        모아 += 전기칸[j]
                        봤다 = True
                    j -= 1
                if j >= 0 and 줄들[j]['깊이'] == 깊 and not 줄들[j]['총계'] \
                        and 전기칸[j] is not None:
                    모아 += 전기칸[j]
                    봤다 = True
                전기칸[i] = 모아 if 봤다 else None
            찾자 = ({'자산총계': 'total assets',
                     '부채및자본총계': 'total liabilities and equity'} if 구분 == 'BS'
                    else {'매출액': 'total income', '매출원가': 'total cost of goods sold',
                          '당기순이익': 'profit'})
            for 라벨, 키 in 찾자.items():
                for x in 줄들:
                    if _맞춤이름(x['이름']) == 키 and x['당기'] is not None:
                        영문합[라벨] = x['당기']
                        break

        # ── 국문 ↔ 영문 대사 (같은 원장에서 나온 숫자인지 확인) ──
        대사칸 = ''
        if 퀵 and 퀵.get('월') and 퀵['월'] != 보고월:
            대사칸 = (
                f'<div class="reconc"><b>국문 ↔ 영문 대사</b> — 기간이 달라 견주지 않았습니다.'
                f'<div class="chks"><span class="chk no">⚠ 화면 자료는 {보고월}월 · '
                f'올려 둔 원장은 {퀵["월"]}월</span></div>'
                f'같은 달로 맞추시려면 「실적보고 엑셀작성」 에서 {퀵["월"]}월 실적 엑셀을 만든 뒤, '
                f'그 파일을 왼쪽 「⚙ 설정 · 파일 업로드」 에 올려 주세요.</div>')
        elif 영문합:
            조각 = []
            for 라벨 in 국문합:
                국, 영 = 국문합.get(라벨), 영문합.get(라벨)
                if 국 is None or 영 is None:
                    continue
                차 = round(국 - 영, 2)
                맞다 = abs(차) < 1.0
                조각.append(
                    f'<span class="chk {"ok" if 맞다 else "no"}">'
                    f'{"✔" if 맞다 else "⚠"} {라벨} · 국문 {국:,.2f} / 영문 {영:,.2f}'
                    f'{"" if 맞다 else f" · 차이 {차:,.2f}"}</span>')
            if 조각:
                모두맞다 = all('chk ok' in s for s in 조각)
                대사칸 = (
                    f'<div class="reconc"><b>국문 ↔ 영문 대사</b> — '
                    f'{"두 표의 합계가 같습니다" if 모두맞다 else "차이가 있습니다"} '
                    f'(원장 「{퀵["시트"]}」 기준)'
                    f'<div class="chks">{"".join(조각)}</div></div>')

        if 국문으로:
            칸들 = ''
            for i, (번호, 국, 영, 수준, 더할) in enumerate(서식):
                이름 = (국 if 국문으로 else 영) or (영 if 국문으로 else 국)
                딴이름 = (영 if 국문으로 else 국)
                if not 영 and not 더할 and 국 in ('자산', '부채', '자본'):
                    칸들 += (f'<tr class="total"><td class="lft" colspan="8" '
                             f'style="background:{T["panel2"]}">{이름}</td></tr>')
                    continue
                전 = _금액(i, '전기')
                당 = _금액(i, '당기')
                굵게 = ' class="total"' if 수준 <= 1 else (' class="sub"' if 수준 == 2 else '')
                들여 = 0 if 수준 == 0 else (12 if 수준 == 1 else (24 if 수준 == 2 else 34))
                칸들 += (
                    f'<tr{굵게}><td class="lft" style="padding-left:{들여 + 11}px">'
                    f'{("<b>" + 번호 + "</b> " if 번호 else "")}{이름}'
                    f'<span style="color:{T["ink3"]}; font-size:12px"> · {딴이름}</span></td>'
                    f'<td>{_돈(전)}</td><td>{_돈(전 * 전기환율, 0)}</td>'
                    f'<td>{_돈(당)}</td><td></td><td></td>'
                    f'<td>{_돈(당)}</td><td>{_돈(당 * 당기환율, 0)}</td></tr>')
            제목 = '재 무 상 태 표' if 구분 == 'BS' else '손 익 계 산 서'
            기간 = (f'제 (당) 기 : {당해연도}년 {보고월:02d}월말 · '
                    f'제 (전) 기 : {당해연도 - 1}년 12월 31일' if 구분 == 'BS'
                    else f'제 (당) 기 : {당해연도}.01.01~{당해연도}.{보고월:02d} · '
                         f'제 (전) 기 : {당해연도 - 1}.01.01~{당해연도 - 1}.12.31')
            맞춤말 = ('' if not 맞춘값 else
                      f'<br><b>미처분이익잉여금</b> 은 원장에 없는 「맞추는 숫자」라 '
                      f'자산총계와 부채및자본총계가 맞도록 채웠습니다 '
                      f'(전기 {_돈(맞춘값.get("전기", 0))} · 당기 {_돈(맞춘값.get("당기", 0))}).')
            st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>{제목} <span class="unitbadge">국문</span></h3>
        <div class="sub">{기간} &nbsp;·&nbsp; 연결패키지 {구분} 시트와 같은 모양입니다</div>
        {대사칸}
        <div class="stick"><table class="lined" style="min-width:1080px; margin-top:8px">
          <colgroup><col style="width:26%"><col style="width:11%"><col style="width:12%">
            <col style="width:11%"><col style="width:9%"><col style="width:9%">
            <col style="width:11%"><col style="width:11%"></colgroup>
          <thead>
            <tr><th>과목</th>
                <th>전기 ({당해연도 - 1}년)<br><span class="th2">현지통화(USD)</span></th>
                <th>전기 ({당해연도 - 1}년)<br><span class="th2">KRW</span></th>
                <th>당기 ({당해연도 % 100}년-수정전)<br><span class="th2">현지통화(USD)</span></th>
                <th>수정분개<br><span class="th2">Debit</span></th>
                <th>수정분개<br><span class="th2">Credit</span></th>
                <th>당기 ({당해연도 % 100}년-수정후)<br><span class="th2">현지통화(USD)</span></th>
                <th>당기 ({당해연도 % 100}년-수정후)<br><span class="th2">KRW</span></th></tr>
          </thead>
          <tbody>{칸들}</tbody></table></div>
        <div class="calcnote"><b>어디서 온 숫자인가</b> —
          전기는 올려 주신 실적 엑셀의
          <b>{'「26년 BS기초」' if 구분 == 'BS' else '「25년 원장_raw」'}</b>,
          당기는 <b>「26년 원장_raw」</b> 에서 계정별로 계산했습니다.
          <b>수정분개</b> 는 연결 결산조정 칸이라 아직 비어 있습니다.{맞춤말}</div>
      </div>
    </div>""")
            그렸다 = True
        elif 퀵:
            # ── 영문 : 퀵북 표를 차례·들여쓰기 그대로 옮겨 놓습니다 ──
            줄들 = 퀵['줄']
            줄 = ''
            for i, x in enumerate(줄들):
                전, 당 = 전기칸[i], x['당기']
                굵게 = ' class="total"' if x['총계'] else (
                    ' class="sub"' if 당 is not None else ' class="head"')
                들여 = 11 + x['깊이'] * 16
                줄 += (f'<tr{굵게}><td class="lft" style="padding-left:{들여}px">'
                       f'{x["이름"]}</td>'
                       f'<td>{"" if 전 is None else _돈(전)}</td>'
                       f'<td>{"" if 당 is None else _돈(당)}</td>'
                       f'<td>{"" if (전 is None or 당 is None) else _돈(당 - 전)}</td></tr>')
            머리 = 퀵['제목']
            st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>{머리[1] or ('Balance Sheet' if 구분 == 'BS'
                                         else 'Profit and Loss')}
          <span class="unitbadge">English</span></h3>
        <div class="sub">{머리[0]} &nbsp;·&nbsp; {머리[2]}
          &nbsp;·&nbsp; prior year {당해연도 - 1} &nbsp;·&nbsp; USD</div>
        {대사칸}
        <div class="stick"><table class="lined" style="min-width:760px; margin-top:8px">
          <colgroup><col style="width:46%"><col style="width:18%">
                    <col style="width:18%"><col style="width:18%"></colgroup>
          <thead><tr><th>Account</th><th>Prior year ({당해연도 - 1})</th>
                     <th>Current ({당해연도})</th><th>Change</th></tr></thead>
          <tbody>{줄}</tbody></table></div>
        <div class="calcnote"><b>어디서 온 표인가</b> — 올려 두신 원장의
          <b>「{퀵['시트']}」</b> 시트를 줄 차례·들여쓰기 그대로 옮겼습니다
          (구역 머리글과 Total 줄 포함 {len(줄들)}줄).
          전기는
          {'실적 엑셀의 「26년 BS기초」' if 구분 == 'BS' else '「25년 원장_raw」 계정별 누계'}
          에서 같은 계정 이름을 찾아 넣었고, Total 줄은 안쪽 줄을 더해 채웠습니다.</div>
      </div>
    </div>""")
            그렸다 = True
        else:
            st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>영문 재무제표 <span class="unitbadge">English</span></h3>
        <div class="calcnote"><b>원장 원본이 아직 없습니다</b> —
          왼쪽 메뉴 <b>「실적보고 엑셀작성」</b> 에서 ① 원장 원본(퀵북 · 그 달 BS·IS 시트가
          들어 있는 파일)을 한 번 올려 주시면, 그 안의
          <b>「{당해연도 % 100}년 {보고월}월(BS)」·「{당해연도 % 100}년 {보고월}월(IS)」</b>
          시트를 차례 그대로 여기에 보여 드립니다.</div>
      </div>
    </div>""")
            그렸다 = True

    # 연결패키지 BS·IS 시트가 쓰는 계정과목 이름 (국문 · 영문)
    미리보기 = {
        'BS': [('Ⅰ. 유동자산', 'Current assets'),
               ('　1. 현금및현금성자산', 'Cash and cash equivalents'),
               ('　2. 매출채권', 'Trade Receivable, gross'),
               ('　3. 재고자산', 'Inventory'),
               ('Ⅱ. 비유동자산', 'Non-current assets'),
               ('　1. 유형자산', 'Property, plant and equipment'),
               ('　2. 사용권자산', 'Right-of-use assets'),
               ('Ⅲ. 부채', 'Liabilities'),
               ('　1. 매입채무', 'Trade payables'),
               ('　2. 차입금', 'Borrowings'),
               ('Ⅳ. 자본', 'Equity')],
        'IS': [('Ⅰ. 매출액', 'Revenue'),
               ('Ⅱ. 매출원가', 'Cost of sales'),
               ('Ⅲ. 매출총이익', 'Gross profit'),
               ('Ⅳ. 판매비와관리비', 'Selling, general and administrative expenses'),
               ('Ⅴ. 영업이익', 'Operating profit'),
               ('Ⅵ. 영업외손익', 'Other income and expenses'),
               ('Ⅶ. 법인세비용차감전순손익', 'Profit before income tax'),
               ('Ⅷ. 당기순손익', 'Profit for the period')],
        'MC': [('Ⅰ. 재료비', 'Raw materials used'),
               ('Ⅱ. 노무비', 'Labour costs'),
               ('Ⅲ. 경비', 'Manufacturing overhead'),
               ('Ⅳ. 당기총제조비용', 'Total manufacturing costs'),
               ('Ⅴ. 기초재공품재고액', 'Work in progress, beginning'),
               ('Ⅵ. 기말재공품재고액', 'Work in progress, ending'),
               ('Ⅶ. 당기제품제조원가', 'Cost of goods manufactured')],
    }.get(구분)
    if 그렸다:
        미리보기 = None
    if 미리보기:
        칸 = ''.join(
            f'<tr class="sub"><td class="name">{국 if 국문으로 else 영}</td>'
            f'<td class="lft" style="color:{T["ink3"]}">{영 if 국문으로 else 국}</td></tr>'
            for 국, 영 in 미리보기)
        st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="card"><h3>{'계정과목 (국문)' if 국문으로 else 'Line items (English)'}
        <span class="unitbadge">{'국문 기준' if 국문으로 else 'English'}</span></h3>
        <div class="sub">연결패키지 「{구분}」 시트가 쓰는 이름입니다 ·
          숫자는 원장이 연결되면 이 자리에 채워집니다</div>
        <table class="lined" style="margin-top:8px">
          <colgroup><col style="width:44%"><col style="width:56%"></colgroup>
          <thead><tr><th>{'국문' if 국문으로 else 'English'}</th>
                     <th>{'English' if 국문으로 else '국문'}</th></tr></thead>
          <tbody>{칸}</tbody></table></div>
    </div>""")

    줄 = ''.join(
        f'<tr class="sub"><td class="name">{i}</td><td class="lft">{무엇}</td>'
        f'<td class="lft">{자료}</td></tr>'
        for i, (무엇, 자료) in enumerate(할일, 1))
    st.html(f"""{CARD_CSS}<div class="wrap" translate="no">
      <div class="kpi-row" style="--n:3">
        <div class="kpi-card"><div class="kpi-label">기준 분기</div>
          <div class="kpi-value" style="font-size:24px;font-weight:700">{당해연도}년 {(보고월 - 1) // 3 + 1}Q</div>
          <div class="kpi-delta"><span class="muted">원장 1~{보고월}월 기준</span></div></div>
        <div class="kpi-card"><div class="kpi-label">화면</div>
          <div class="kpi-value" style="font-size:24px;font-weight:700">{이름}</div>
          <div class="kpi-delta"><span class="muted">{한줄}</span></div></div>
        <div class="kpi-card"><div class="kpi-label">준비 상태</div>
          <div class="kpi-value" style="font-size:24px;font-weight:700;color:{T['accent']}">준비 중</div>
          <div class="kpi-delta"><span class="muted">연결 분개장을 받으면 채웁니다</span></div></div>
      </div>
      <div class="card"><h3>이 화면에 들어갈 것</h3>
        <div class="sub">쓸 자료가 이미 있는 것부터 순서대로 채워 나갑니다</div>
        <table class="lined">
          <colgroup><col style="width:8%"><col style="width:50%"><col style="width:42%"></colgroup>
          <thead><tr><th>순서</th><th>내용</th><th>쓰는 자료</th></tr></thead>
          <tbody>{줄}</tbody></table></div>
    </div>""")

    받을칸, 여백 = st.columns([1, 3], gap='small')
    with 받을칸:
        st.download_button(f'⤓  {이름} 엑셀 내려받기', data=b'', file_name='준비중.xlsx',
                           disabled=True, key=f'pkg_{구분}', width='stretch',
                           help='연결 분개장을 받아 채우면 여기서 내려받을 수 있게 됩니다')
    with 여백:
        st.caption('※ 검증은 CTKonnect 원장의 BS · IS 기초자료로 진행합니다.')
