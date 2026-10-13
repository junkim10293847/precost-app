"""사전원가 자동입력 — Streamlit 앱
실행: streamlit run app.py
"""
import io, tempfile, urllib.parse
import pandas as pd
import streamlit as st
import openpyxl

from core import (read_item_list, collect_files, run, to_excel, read_owner_master, assign_owners,
                  owner_requests, request_mail, read_postcost, compare_pre_post)

st.set_page_config(page_title='사전원가 자동입력', page_icon='📑', layout='wide')
st.title('📑 사전원가 자동입력')
st.caption('요청 품목 목록 + 견적서(zip) → 원가 자동입력 · 담당자별 요청 메일 · 사전/사후원가 비교')

with st.sidebar:
    st.header('사용 방법')
    st.markdown(
        '1. **요청 품목 목록** 엑셀\n2. **견적서** zip 또는 여러 파일\n'
        '3. (선택) **지난달 결과 파일** → 담당자 자동 배정\n4. (선택) **사후원가 파일** → 사전/사후 비교\n'
        '5. 실행 → 노란 행만 확인 → 다운로드\n\n'
        '**파일명 규칙**: 견적서 파일명 앞에 목록 번호 (예: `29. 테아닌.xlsx`, `82~85.`)')
    period = st.text_input('대상 월', value='2026년 7월')
    sender = st.text_input('발신 부서', value='재무지원팀')
    due = st.text_input('회신 기한(선택)', value='')
    match_mode = st.radio('견적서 매칭 방식', ['파일번호 + 품목명 (권장)', '품목명만 (번호 없는 파일)'], index=0)

c1, c2 = st.columns(2)
with c1:
    list_file = st.file_uploader('① 사전원가 요청 품목 목록 (엑셀)', type=['xlsx', 'xlsm'])
    sheet = None
    if list_file:
        names = openpyxl.load_workbook(io.BytesIO(list_file.getvalue()), read_only=True).sheetnames
        default = next((i for i, n in enumerate(names) if '사전원가' in n or 'List' in n), 0)
        sheet = st.selectbox('목록 시트', names, index=default)
    owner_files = st.file_uploader('③ (선택) 지난달 결과 파일 — 담당자 자동 배정용', type=['xlsx'],
                                   accept_multiple_files=True)
with c2:
    quote_files = st.file_uploader('② 견적서 파일 (zip 또는 여러 파일)', accept_multiple_files=True,
                                   type=['zip', 'xlsx', 'xls', 'xlsm', 'pdf', 'png', 'jpg', 'jpeg'])
    post_file = st.file_uploader('④ (선택) 사후원가 파일 — 제품명·제조번호·원료비·부자재비·노무비/경비·생산량',
                                 type=['xlsx'])
    post_sheet, post_unit = None, None
    if post_file:
        pn = openpyxl.load_workbook(io.BytesIO(post_file.getvalue()), read_only=True).sheetnames
        post_sheet = st.selectbox('사후원가 시트', pn)
        post_unit = {'자동(생산량 열 있으면 금액÷생산량)': None, '이미 단위원가': True,
                     '금액 → 생산량으로 나눔': False}[st.radio('사후원가 금액 형태',
                                                         ['자동(생산량 열 있으면 금액÷생산량)', '이미 단위원가',
                                                          '금액 → 생산량으로 나눔'], horizontal=True)]

if st.button('🚀 자동 입력 실행', type='primary', disabled=not (list_file and quote_files)):
    try:
        items, _ = read_item_list(list_file.getvalue(), sheet)
    except Exception as e:
        st.error(f'목록을 읽지 못했습니다: {e}')
        st.stop()
    files = collect_files([(f.name, f.getvalue()) for f in quote_files], tempfile.mkdtemp())
    bar = st.progress(0.0, text='준비 중…')
    rows = run(items, files, progress=lambda i, n, nm: bar.progress(i / n, text=f'{i}/{n} · {nm[:30]}'),
               use_numbers=match_mode.startswith('파일번호'))
    bar.empty()
    if owner_files:
        assign_owners(rows, read_owner_master([(f.name, f.getvalue()) for f in owner_files]))
    else:
        for r in rows:
            r['owner'] = r.get('owner') or '미지정'
    cmp = None
    if post_file:
        try:
            post, psh, pcols, used_qty = read_postcost(post_file.getvalue(), post_sheet, post_unit)
            cmp = compare_pre_post(rows, post)
            st.session_state['post_info'] = (psh, pcols, used_qty)
        except Exception as e:
            st.warning(f'사후원가 파일을 읽지 못했습니다: {e}')
    st.session_state.update(rows=rows, cmp=cmp)

rows = st.session_state.get('rows')
if not rows:
    st.stop()

n = len(rows)
m = st.columns(5)
m[0].metric('요청 품목', n)
m[1].metric('자동 완료', sum(1 for r in rows if r['status'] == 'O' and not r['review']))
m[2].metric('🟡 검토 필요', sum(1 for r in rows if r['review']))
m[3].metric('⚪ 견적서 없음', sum(1 for r in rows if r['status'] == 'X'))
m[4].metric('📧 요청 대상 담당자', len(owner_requests(rows)))

tab1, tab2, tab3 = st.tabs(['① 자동입력 결과', '② 담당자별 요청 메일', '③ 사전 vs 사후원가'])

with tab1:
    df = pd.DataFrame([{
        '구분': ('⚪' if r['status'] == 'X' else '🟡' if r['review'] else '✅'),
        'no': r['no'], '제품명': r['name'], '거래처': r['customer'], '담당자': r.get('owner'), '단가': r['price'],
        '원재료비': r['raw'], '부원료비': r['sub'], '부재료비': r['pack'], '가공비': r['proc'],
        '마진': r['margin'], '최종견적': r['final'], 'O/X': r['status'],
        '매칭 견적서': r['file'], '메모': ' / '.join(dict.fromkeys(r['notes'])),
    } for r in rows])
    only = st.toggle('검토 필요·견적서 없음만 보기', value=False)
    view = df[df['구분'] != '✅'] if only else df
    st.caption('숫자·담당자를 표에서 직접 고치면 다운로드 엑셀과 요청 메일에 반영됩니다.')
    edited = st.data_editor(
        view, use_container_width=True, height=520, hide_index=True,
        disabled=['구분', 'no', '제품명', '거래처', '단가', '매칭 견적서'],
        column_config={k: st.column_config.NumberColumn(format='%d') for k in
                       ['단가', '원재료비', '부원료비', '부재료비', '가공비', '마진', '최종견적']})
    keymap = {'원재료비': 'raw', '부원료비': 'sub', '부재료비': 'pack', '가공비': 'proc', '마진': 'margin',
              '최종견적': 'final', 'O/X': 'status', '담당자': 'owner'}
    for idx, er in edited.iterrows():
        r = rows[idx]
        for col, k in keymap.items():
            v = er[col]
            r[k] = None if (v is None or (isinstance(v, float) and pd.isna(v))) else v
        r['notes'] = [x for x in str(er['메모']).split(' / ') if x and x != 'nan']

with tab2:
    reqs = owner_requests(rows)
    if not reqs:
        st.success('요청할 품목이 없습니다.')
    st.caption('담당자 이메일을 넣으면 [메일 쓰기] 버튼으로 기본 메일 프로그램이 열립니다. (자동 발송은 하지 않음)')
    emails = st.session_state.setdefault('emails', {})
    for owner, lst in reqs.items():
        subj, body = request_mail(owner, lst, period, sender, due or None)
        with st.expander(f'{owner} — {len(lst)}건', expanded=False):
            emails[owner] = st.text_input('이메일', value=emails.get(owner, ''), key=f'em_{owner}')
            st.text_input('제목', value=subj, key=f'sj_{owner}')
            st.text_area('본문', value=body, height=260, key=f'bd_{owner}')
            if emails[owner]:
                url = f"mailto:{emails[owner]}?subject={urllib.parse.quote(subj)}&body={urllib.parse.quote(body)}"
                st.link_button('✉️ 메일 쓰기', url)

with tab3:
    cmp = st.session_state.get('cmp')
    if not cmp:
        st.info('④ 사후원가 파일을 올리면 품목별 사전(견적) vs 사후(실적) 단위원가 차이를 계산합니다.')
    else:
        psh, pcols, used_qty = st.session_state['post_info']
        st.caption(f"시트 '{psh}' · 생산량 나눔 {'예' if used_qty else '아니오'} · "
                   f"원재료={pcols['raw']} / 부재료={pcols['pack']} / 가공비={pcols['proc']}")
        cdf = pd.DataFrame([{
            'no': c['no'], '제품명': c['name'], '거래처': c['customer'], '담당자': c.get('owner'),
            '최종견적': c['final'], '사전원가': c['pre_cost'], '사후원가': c.get('post_cost'),
            '차이': c.get('diff'), '차이율%': None if c.get('diff_rate') is None else c['diff_rate'] * 100,
            '사전이익률%': None if c.get('pre_margin_rate') is None else c['pre_margin_rate'] * 100,
            '사후이익률%': None if c.get('post_margin_rate') is None else c['post_margin_rate'] * 100,
            '매칭': c['match']} for c in cmp])
        k = st.columns(3)
        k[0].metric('사후원가 매칭', f"{(cdf['사후원가'].notna()).sum()} / {len(cdf)}")
        k[1].metric('원가 5% 초과 상승', int((cdf['차이율%'] > 5).sum()))
        k[2].metric('사후 적자 품목', int((cdf['사후이익률%'] < 0).sum()))
        st.dataframe(cdf.sort_values('차이율%', ascending=False, na_position='last'), use_container_width=True,
                     hide_index=True, height=480,
                     column_config={c: st.column_config.NumberColumn(format='%,.0f') for c in
                                    ['최종견적', '사전원가', '사후원가', '차이']} |
                                   {c: st.column_config.NumberColumn(format='%.1f') for c in
                                    ['차이율%', '사전이익률%', '사후이익률%']})

st.divider()
title = f'{period} 사전원가 요청 품목 목록 (자동입력)'
st.download_button('📥 엑셀 다운로드 (->영업요청 · 담당자별 요청 · 사전vs사후)',
                   data=to_excel(rows, title, compare=st.session_state.get('cmp'), period=period),
                   file_name=f'{title}.xlsx', type='primary',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
