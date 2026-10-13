"""사전원가 자동입력 — Streamlit 앱
실행: streamlit run app.py
로그인 계정은 Streamlit Cloud → App settings → Secrets 에 등록 (README 참고)
"""
import io, os, hmac, hashlib, tempfile, urllib.parse
import pandas as pd
import streamlit as st
import openpyxl

from core import (preview_quote, read_item_list, collect_files, run, to_excel, read_owner_master, assign_owners,
                  owner_requests, request_mail, read_postcost, compare_pre_post, reselect_file, refresh_review)

st.set_page_config(page_title='사전원가 자동입력', page_icon='📑', layout='wide')


# ================================================================ 로그인
def _users():
    try:
        return dict(st.secrets.get('users', {}))
    except Exception:
        return {}


def _check(pw_input, stored):
    stored = str(stored)
    if stored.startswith('sha256:'):
        return hmac.compare_digest(hashlib.sha256(pw_input.encode()).hexdigest(), stored[7:])
    return hmac.compare_digest(pw_input, stored)


def login_gate():
    if st.session_state.get('user'):
        return
    users = _users()
    st.title('📑 사전원가 자동입력')
    if not users:
        st.error('로그인 계정이 아직 등록되지 않았습니다. 관리자: Streamlit Cloud → Manage app → Settings → '
                 'Secrets 에 [users] 를 등록하세요.')
        st.code('[users]\nadmin = "비밀번호"\nfinance1 = "비밀번호"', language='toml')
        st.stop()
    with st.form('login'):
        uid = st.text_input('아이디')
        pw = st.text_input('비밀번호', type='password')
        ok = st.form_submit_button('로그인', type='primary')
    if ok:
        if uid in users and _check(pw, users[uid]):
            st.session_state['user'] = uid
            st.rerun()
        else:
            st.error('아이디 또는 비밀번호가 맞지 않습니다.')
    st.stop()


login_gate()

# ================================================================ 화면
st.title('📑 사전원가 자동입력')
st.caption('요청 품목 목록 + 견적서(zip) → 원가 자동입력 · 화면에서 수정/확인 후 저장 · 담당자별 요청 메일 · 사전/사후원가 비교')

with st.sidebar:
    st.markdown(f"👤 **{st.session_state['user']}** 님")
    if st.button('로그아웃'):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
    st.divider()
    st.header('사용 방법')
    st.markdown(
        '1. **요청 품목 목록** 엑셀\n2. **견적서** zip 또는 여러 파일\n'
        '3. (선택) **지난달 결과 파일** → 담당자 자동 배정\n4. (선택) **사후원가 파일** → 사전/사후 비교\n'
        '5. 실행 → 노란 행 수정·**확인 체크** → **💾 저장** → 다운로드\n\n'
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
        opt = st.radio('사후원가 금액 형태', ['자동(생산량 열 있으면 금액÷생산량)', '이미 단위원가', '금액 → 생산량으로 나눔'],
                       horizontal=True)
        post_unit = {'자동(생산량 열 있으면 금액÷생산량)': None, '이미 단위원가': True, '금액 → 생산량으로 나눔': False}[opt]

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
    for r in rows:
        r['owner'] = r.get('owner') or '미지정'
    post = None
    if post_file:
        try:
            post, psh, pcols, used_qty = read_postcost(post_file.getvalue(), post_sheet, post_unit)
            st.session_state['post_info'] = (psh, pcols, used_qty)
        except Exception as e:
            st.warning(f'사후원가 파일을 읽지 못했습니다: {e}')
    st.session_state.update(rows=rows, files=files, post=post, cache={}, ver=0, saved_msg=None)

rows = st.session_state.get('rows')
if not rows:
    st.stop()
# 이전 버전 화면에서 넘어온 결과 등 필수 정보가 없으면 다시 실행하도록 안내 (버전 교체 시 오류 방지)
if 'files' not in st.session_state or any('edits' not in r for r in rows):
    st.session_state.pop('rows', None)
    st.info('앱이 새 버전으로 바뀌었습니다. 위에서 파일을 확인하고 🚀 자동 입력 실행을 다시 눌러 주세요.')
    st.stop()
st.session_state.setdefault('ver', 0)
st.session_state.setdefault('cache', {})
st.session_state.setdefault('post', None)

files = st.session_state['files']
path_by_name = {os.path.basename(p): p for p in files}

n = len(rows)
m = st.columns(5)
m[0].metric('요청 품목', n)
m[1].metric('✅ 완료(자동+확인)', sum(1 for r in rows if r['status'] == 'O' and not r['review']))
m[2].metric('🟡 검토 필요', sum(1 for r in rows if r['review']))
m[3].metric('⚪ 견적서 없음', sum(1 for r in rows if r['status'] == 'X' and not r.get('checked')))
m[4].metric('✔ 확인완료 체크', sum(1 for r in rows if r.get('checked')))

tab0, tab1, tab2, tab3 = st.tabs(['🔍 한 건씩 검토 (견적서 비교)', '📋 전체 표 · 일괄 수정', '✉️ 담당자별 요청 메일', '📊 사전 vs 사후원가'])

NUM = ['원재료비', '부원료비', '부재료비', '가공비', '마진', '최종견적']
KEY = {'원재료비': 'raw', '부원료비': 'sub', '부재료비': 'pack', '가공비': 'proc', '마진': 'margin', '최종견적': 'final'}


def _ival(v):
    return None if v is None or (isinstance(v, float) and pd.isna(v)) else int(v)


with tab0:
    if st.session_state.get('saved_msg'):
        st.success(st.session_state['saved_msg'])
    flt = st.radio('대상', ['🟡 검토 필요', '⚪ 견적서 없음', '✔ 확인완료', '전체'], horizontal=True, key='rv_filter')
    cond = {'🟡 검토 필요': lambda r: r['review'],
            '⚪ 견적서 없음': lambda r: r['status'] == 'X' and not r.get('checked'),
            '✔ 확인완료': lambda r: r.get('checked'),
            '전체': lambda r: True}[flt]
    idxs = [i for i, r in enumerate(rows) if cond(r)]
    if not idxs:
        st.success('해당하는 품목이 없습니다. 👍')
    else:
        pos = st.session_state.get('rv_pos', 0)
        pos = min(pos, len(idxs) - 1)
        labels = [f"[{rows[i]['no']}] {rows[i]['name'][:40]}" for i in idxs]
        nav1, nav2, nav3 = st.columns([1, 6, 1])
        if nav1.button('◀ 이전', disabled=pos == 0, use_container_width=True):
            st.session_state['rv_pos'] = pos - 1
            st.rerun()
        sel = nav2.selectbox(f'품목 ({pos + 1} / {len(idxs)})', range(len(idxs)), index=pos,
                             format_func=lambda k: labels[k], label_visibility='collapsed')
        if sel != pos:
            st.session_state['rv_pos'] = sel
            st.rerun()
        if nav3.button('다음 ▶', disabled=pos >= len(idxs) - 1, use_container_width=True):
            st.session_state['rv_pos'] = pos + 1
            st.rerun()

        i = idxs[pos]
        r = rows[i]
        left, right = st.columns([2, 3], gap='large')
        with left:
            st.markdown(f"#### [{r['no']}] {r['name']}")
            st.caption(f"거래처 **{r['customer'] or '-'}** · LOT {r['lot'] or '-'} · 규격 {r['spec'] or '-'} · "
                       f"단가 **{(r['price'] or 0):,.0f}원** · 담당자 {r.get('owner') or '-'}")
            for note in dict.fromkeys(r.get('edits', []) + r['notes']):
                (st.warning if ('확인필요' in note or '미추출' in note or 'OCR' in note or '불일치' in note)
                 else st.info)(note, icon='📝')
            ver = st.session_state['ver']
            with st.form(f'rv_{i}_{ver}'):
                c_a, c_b = st.columns(2)
                vals = {}
                for j, col in enumerate(NUM):
                    vals[col] = (c_a if j % 2 == 0 else c_b).number_input(
                        col, value=_ival(r[KEY[col]]), step=1, format='%d', key=f'rv_{col}_{i}_{ver}')
                cost = sum(v or 0 for k, v in vals.items() if k not in ('마진', '최종견적'))
                st.caption(f"제조원가 {cost:,.0f} + 마진 {(vals['마진'] or 0):,.0f} = {cost + (vals['마진'] or 0):,.0f}  "
                           f"│ 최종견적 {(vals['최종견적'] or 0):,.0f}  │ 단가-견적 "
                           f"{((r['price'] or 0) - (vals['최종견적'] or 0)):+,.0f}")
                owner = st.text_input('담당자', value=r.get('owner') or '', key=f'rv_own_{i}_{ver}')
                ox = st.radio('O/X', ['O', 'X'], index=0 if r['status'] == 'O' else 1, horizontal=True,
                              key=f'rv_ox_{i}_{ver}')
                chk = st.checkbox('✔ 견적서와 대조해 확인 완료', value=bool(r.get('checked')), key=f'rv_chk_{i}_{ver}')
                b1, b2 = st.columns(2)
                save_one = b1.form_submit_button('💾 저장', use_container_width=True)
                save_next = b2.form_submit_button('💾 저장 후 다음 ▶', type='primary', use_container_width=True)
            if save_one or save_next:
                ch = []
                for col in NUM:
                    nv, ov = _ival(vals[col]), _ival(r[KEY[col]])
                    if nv != ov:
                        r[KEY[col]] = nv
                        ch.append(f"{col} {ov if ov is not None else '-'}→{nv if nv is not None else '-'}")
                if ox != r['status']:
                    ch.append(f"O/X {r['status']}→{ox}")
                    r['status'] = ox
                if owner.strip() and owner.strip() != (r.get('owner') or ''):
                    r['owner'], r['owner_src'] = owner.strip(), '수동'
                    ch.append(f'담당자→{owner.strip()}')
                r.setdefault('edits', []).extend(ch)
                r['checked'] = chk
                refresh_review(r)
                st.session_state['ver'] += 1
                st.session_state['saved_msg'] = f"[{r['no']}] 저장 완료" + (f" — {', '.join(ch)}" if ch else '') + \
                    (' · 확인완료' if chk else '')
                if save_next:
                    # 확인완료로 목록에서 빠지면 같은 위치가 곧 다음 품목
                    still = cond(r)
                    st.session_state['rv_pos'] = pos + 1 if still else pos
                st.rerun()

            with st.expander('📎 다른 견적서 파일로 바꾸기'):
                cur = r['file'] or ''
                opts = [''] + sorted(path_by_name)
                newf = st.selectbox('견적서 파일', opts, index=opts.index(cur) if cur in opts else 0,
                                    key=f'rv_file_{i}_{st.session_state["ver"]}')
                if st.button('이 파일로 다시 읽기', disabled=newf == cur, key=f'rv_refile_{i}'):
                    rows[i] = reselect_file(r, path_by_name.get(newf), st.session_state['cache'])
                    st.session_state['ver'] += 1
                    st.session_state['saved_msg'] = f"[{r['no']}] 견적서를 '{newf or '없음'}'로 바꿔 다시 읽었습니다."
                    st.rerun()

        with right:
            fpath = path_by_name.get(r['file']) if r['file'] else None
            st.markdown(f"#### 📄 견적서 원본")
            if not fpath:
                st.info('매칭된 견적서가 없습니다. 왼쪽 "다른 견적서 파일로 바꾸기"에서 고를 수 있습니다.')
            else:
                st.caption(r['file'])
                with open(fpath, 'rb') as fh:
                    st.download_button('⬇ 원본 파일 열기(다운로드)', fh.read(), file_name=r['file'], key=f'dl_{i}')
                kind, data = preview_quote(fpath, r.get('sheet'))
                if kind == 'images':
                    zoom = st.slider('확대', 50, 200, 100, 10, key='rv_zoom', format='%d%%')
                    for im in data:
                        st.image(im, width=int(700 * zoom / 100))
                elif kind == 'tables':
                    names = [t for t, _ in data]
                    shn = st.selectbox('시트', names, index=0, key=f'rv_sheet_{i}')
                    df_s = dict(data)[shn]
                    kw = ('합계', '합 계', '원가', '마진', '이익', '출고', '견적가', '원재료', '부재료', '가공', '자재비')
                    st.dataframe(df_s.style.apply(
                        lambda row: ['background-color:#FFF2CC' if any(k in ' '.join(row.values) for k in kw) else ''
                                     for _ in row], axis=1),
                        use_container_width=True, height=620)
                    st.caption('노란 줄 = 합계·원가·마진·출고가 등 금액 줄')
                else:
                    st.warning(data)

with tab1:
    if st.session_state.get('saved_msg'):
        st.success(st.session_state['saved_msg'])
    df = pd.DataFrame([{
        '확인': bool(r.get('checked')),
        '구분': ('✔' if r.get('checked') else '⚪' if r['status'] == 'X' else '🟡' if r['review'] else '✅'),
        'no': r['no'], '제품명': r['name'], '거래처': r['customer'], '담당자': r.get('owner'), '단가': r['price'],
        '원재료비': r['raw'], '부원료비': r['sub'], '부재료비': r['pack'], '가공비': r['proc'],
        '마진': r['margin'], '최종견적': r['final'], 'O/X': r['status'],
        '매칭 견적서': r['file'] or '',
        '메모': ' / '.join(dict.fromkeys(r.get('edits', []) + r['notes'])),
    } for r in rows])

    f1, f2 = st.columns([1, 3])
    only = f1.toggle('검토 필요·견적서 없음만 보기', value=False)
    f2.caption('① 숫자·담당자·O/X를 고치거나 ② **매칭 견적서**를 다른 파일로 바꾸고(그 파일로 다시 읽음) '
               '③ 맞으면 **확인** 체크 → 반드시 **💾 저장**을 눌러야 아래 엑셀 다운로드에 반영됩니다. '
               '표 오른쪽 위 ⛶ 버튼을 누르면 전체 화면으로 볼 수 있습니다. 한 건씩 비교하려면 🔍 탭을 쓰세요.')
    view = df[df['구분'].isin(['🟡', '⚪'])] if only else df

    with st.form(f"edit_{st.session_state['ver']}"):
        edited = st.data_editor(
            view, use_container_width=True, height=540, hide_index=True,
            disabled=['구분', 'no', '제품명', '거래처', '단가', '메모'],
            column_config={
                '확인': st.column_config.CheckboxColumn('확인', help='내용을 확인했으면 체크 → 노란색 해제', width='small'),
                'O/X': st.column_config.SelectboxColumn('O/X', options=['O', 'X'], width='small'),
                '매칭 견적서': st.column_config.SelectboxColumn(
                    '매칭 견적서', options=[''] + sorted(path_by_name), width='large',
                    help='다른 견적서를 고르면 저장 시 그 파일로 다시 추출합니다'),
                **{k: st.column_config.NumberColumn(k, format='%d', step=1, width='small') for k in ['단가'] + NUM},
                'no': st.column_config.NumberColumn('no', width='small'),
                '구분': st.column_config.TextColumn('구분', width='small'),
                '메모': st.column_config.TextColumn('메모', width='medium'),
            })
        save = st.form_submit_button('💾 저장 (수정·확인 내용 반영)', type='primary')

    if save:
        cache = st.session_state['cache']
        n_edit = n_file = n_chk = 0
        for idx, er in edited.iterrows():
            r = rows[idx]
            orig = df.loc[idx]
            # 1) 견적서 파일 변경 → 재추출
            new_file = er['매칭 견적서'] or ''
            if new_file != (orig['매칭 견적서'] or ''):
                r = reselect_file(r, path_by_name.get(new_file), cache)
                rows[idx] = r
                n_file += 1
            else:
                # 2) 숫자 수정 (원래 값과 다를 때만, 변경 이력 기록)
                for col in NUM:
                    nv, ov = er[col], orig[col]
                    nv = None if pd.isna(nv) else int(nv)
                    ov = None if pd.isna(ov) else int(ov)
                    if nv != ov:
                        r[KEY[col]] = nv
                        r.setdefault('edits', []).append(f"{col} {ov if ov is not None else '-'}→{nv if nv is not None else '-'}")
                        n_edit += 1
                if er['O/X'] != orig['O/X']:
                    r['status'] = er['O/X']
                    r.setdefault('edits', []).append(f"O/X {orig['O/X']}→{er['O/X']}")
                    n_edit += 1
            if (er['담당자'] or '') != (orig['담당자'] or ''):
                r['owner'] = er['담당자'] or '미지정'
                r['owner_src'] = '수동'
                n_edit += 1
            if bool(er['확인']) != bool(orig['확인']):
                r['checked'] = bool(er['확인'])
                n_chk += 1
            refresh_review(r)
        st.session_state['ver'] += 1
        st.session_state['saved_msg'] = (f'저장 완료 — 값 수정 {n_edit}건 · 견적서 변경 {n_file}건 · 확인 체크 변경 {n_chk}건. '
                                         f'엑셀 다운로드에 반영되었습니다.')
        st.rerun()

with tab2:
    reqs = owner_requests(rows)
    if not reqs:
        st.success('요청할 품목이 없습니다.')
    st.caption('확인 체크한 품목은 요청 대상에서 빠집니다. 이메일을 넣으면 [메일 쓰기]로 메일 프로그램이 열립니다(자동 발송 없음).')
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

post = st.session_state.get('post')
cmp = compare_pre_post(rows, post) if post else None
with tab3:
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
                     hide_index=True, height=480)

st.divider()
title = f'{period} 사전원가 요청 품목 목록 (자동입력)'
st.caption('💾 저장한 내용까지 반영됩니다. 표에서 고친 뒤 저장을 누르지 않으면 반영되지 않습니다.')
st.download_button('📥 엑셀 다운로드 (->영업요청 · 담당자별 요청 · 사전vs사후)',
                   data=to_excel(rows, title, compare=cmp, period=period),
                   file_name=f'{title}.xlsx', type='primary',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
