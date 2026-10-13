"""견적서(엑셀/PDF/이미지)에서 원재료비·부재료비·가공비·마진·최종견적을 추출하는 엔진."""
import re, os, subprocess
import openpyxl

NUM_RE = re.compile(r'^-?\(?[\d,]+(\.\d+)?\)?$')


def to_num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace('원', '').strip()
        if NUM_RE.match(s) and any(ch.isdigit() for ch in s):
            try:
                return float(s.replace(',', '').replace('(', '-').replace(')', ''))
            except ValueError:
                return None
    return None


def norm(s):
    return re.sub(r'[\s\*\.\d\(\)（）]', '', str(s))


# ---------- 파일 → 행(row) 토큰 리스트 ----------
def rows_from_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        rows = []
        for r in ws.iter_rows():
            toks = [c.value for c in r if c.value is not None and str(c.value).strip() != '']
            if toks:
                rows.append(toks)
        sheets.append((ws.title, rows))
    return sheets


KEYWORDS = ['합계', '제조원가', '마진', '출고가', '이익', '원재료비', '부재료비', '부원료비', '가공비',
            '임가공비', '최종견적가', '원자재', '부자재', '노무비', '제조경비', '일반경비', '일반비용']


def squeeze(line):
    for k in sorted(KEYWORDS, key=len, reverse=True):
        line = re.sub(r'\s*'.join(k), k, line)
    return line


def rows_from_text(text):
    rows = []
    for line in text.splitlines():
        line = squeeze(line)
        if not line.strip():
            continue
        parts = [p for p in re.split(r'\s{2,}', line.strip()) if p]
        toks = []
        for p in parts:
            # "4,609.4  1,102,500" 같이 공백 하나로 붙은 숫자 분리
            sub = p.split(' ')
            if len(sub) > 1 and all(to_num(x) is not None for x in sub):
                toks.extend(sub)
            else:
                toks.append(p)
        rows.append(toks)
    return [('pdf', rows)]


def rows_from_pdf(path):
    txt = subprocess.run(['pdftotext', '-layout', path, '-'], capture_output=True, text=True).stdout
    ocr = False
    if len(re.sub(r'\s', '', txt)) < 50:  # 스캔본 → OCR
        txt, ocr = ocr_file(path), True
    return rows_from_text(txt), ocr


def ocr_file(path):
    import pytesseract
    from PIL import Image
    imgs = []
    if path.lower().endswith('.pdf'):
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            imgs = [p.to_image(resolution=300).original for p in pdf.pages]
    else:
        imgs = [Image.open(path)]
    langs = 'kor+eng' if 'kor' in pytesseract.get_languages() else 'eng'
    return '\n'.join(pytesseract.image_to_string(im, lang=langs, config='--psm 6') for im in imgs)


# ---------- 행 → 값 추출 ----------
LABELS = {
    'stated': ['확정가', '제안가격', '최종견적가', '최종견적', '견적가', '공급가', '조정가'],
    'raw': ['원재료비', '원료비', '원자재'],
    'sub': ['부원료비'],
    'pack': ['부재료비', '부자재비', '자재비', '부자재', '제료비'],
    'proc': ['가공비', '임가공비'],
    'cost': ['제조원가'],
    'margin': ['이익', '마진', '이윤'],
    'final': ['출고가'],
}
FIELDS = ['raw', 'pack', 'proc', 'margin', 'final']


def nums(toks):
    return [to_num(t) for t in toks if to_num(t) is not None]


def label_of(text):
    n = norm(text)
    for key, labs in LABELS.items():
        for lab in labs:
            if n.startswith(lab):
                return key
    return None


def is_round(v):
    return v is not None and v > 0 and abs(v - round(v / 10) * 10) < 1e-6


def extract_from_rows(rows):
    lab, src = {}, {}
    margin_row_seen = False
    stated = None
    # 1) 명시 라벨
    for toks in rows:
        for j, t in enumerate(toks):
            if not isinstance(t, str) or to_num(t) is not None:
                continue
            key = label_of(t)
            # 문장 안에 들어있는 '조정가 3,029원', '견적가 1750원', '공급가 8,590원'
            m = re.search(r'(확정가|조정가|견적가|공급가|제안가격|최종견적가?)[^\d]{0,15}([\d,]{3,})\s*원?', t.replace(' ', ''))
            if m and stated is None:
                stated = to_num(m.group(2))
            if not key:
                continue
            after = nums(toks[j + 1:])
            if key == 'margin':
                margin_row_seen = True
            if key == 'stated':
                if stated is None and after:
                    stated = after[0]
                continue
            if key == 'final':
                inline = re.findall(r'(?<![\d.])(\d{1,3}(?:,\d{3})+|\d{3,})\s*원', t)
                if inline:
                    stated = stated or to_num(inline[-1])
                elif key in lab and after and is_round(after[0]) and after[0] != lab[key]:
                    stated = stated or after[0]   # 두 번째 '출고가: 10,300' = 확정가
            if key in lab:
                continue
            nxt = toks[j + 1] if j + 1 < len(toks) else None
            if isinstance(nxt, str) and to_num(nxt) is None and label_of(nxt) is None:
                continue  # 섹션 헤더/병합셀 어긋남 (예: '마진 | 합 계 | 410')
            if key in ('raw', 'pack', 'proc'):
                if len(after) > 2:
                    continue
            if key == 'final':
                inline = re.findall(r'(?<![\d.])(\d{1,3}(?:,\d{3})+|\d{3,})\s*원', t)
                if inline:
                    stated = stated or to_num(inline[-1])
                if len(after) >= 2 and is_round(after[-1]) and after[-1] <= after[0]:
                    stated = stated or after[-1]   # 계산값 옆에 확정가 기재
            if after:
                lab[key] = after[0]
                src[key] = 'label-num' if re.match(r'^\s*\d+\s*[.)]', t) else 'label'
    # 2) 가로형 헤더 (원자재 | 부자재 | 임가공비 | 제조원가 | 이익 | 합계)
    for i, toks in enumerate(rows):
        keys = [label_of(t) if isinstance(t, str) else None for t in toks]
        found = [k for k in keys if k]
        if len(set(found)) >= 3 and not nums(toks):
            tail_sum = isinstance(toks[-1], str) and norm(toks[-1]).startswith('합계')
            for k2 in range(i + 1, min(i + 4, len(rows))):
                ns = nums(rows[k2])
                if len(ns) >= len(found):
                    for k, v in zip(found + (['final'] if tail_sum else []), ns):
                        lab[k] = v; src[k] = 'header-row'
                    break
    # 3) 합계 행 순서 (원료 합계 → 자재 합계 → 임가공 합계)
    sums = []
    for toks in rows:
        heads = [norm(str(t)) for t in toks[:3] if isinstance(t, str)]
        if any(h.startswith('합계') for h in heads):
            ns = nums(toks)
            if ns:
                sums.append(ns)
    is_raw = lambda ns: len(ns) >= 3 and 0 <= ns[0] <= 101.5
    cands = [k for k, ns in enumerate(sums) if is_raw(ns)]
    seq = {}
    if cands:
        nz = [k for k in cands if sums[k][-1] != 0]
        ri = nz[0] if nz else cands[0]
        cap = stated or lab.get('final') or lab.get('cost')
        tail = sums[ri][2:]
        pick = [v for v in tail if cap is None or v < cap]
        seq['raw'] = (pick or tail)[0]
        rest = [s for s in sums[ri + 1:] if not is_raw(s)]
        if len(rest) >= 1:
            seq['pack'] = rest[0][0]
        if len(rest) >= 2:
            seq['proc'] = rest[1][0]
    res = {}
    for k in ('raw', 'pack', 'proc'):
        if k in seq and (k not in lab or src.get(k) == 'label'):
            # 번호 매긴 요약표(1. 원재료비 …)·가로형 요약이 없으면 합계 순서값 채택
            res[k] = seq[k]; src[k] = 'sum-seq'
        elif k in lab:
            res[k] = lab[k]
    for k in ('cost', 'margin', 'final', 'sub'):
        if k in lab:
            res[k] = lab[k]
    if 'margin' not in res and (margin_row_seen or 'final' in res):
        res['margin'] = 0.0; src['margin'] = 'blank→0'
    computed = res.get('final')
    if stated:
        res['final'] = stated; src['final'] = 'stated'
    elif computed is not None:
        src['final'] = 'computed'
    res['_computed_final'] = computed
    return res, src


def score(res):
    return sum(1 for k in FIELDS if res.get(k) is not None)


def extract_file(path):
    ext = os.path.splitext(path)[1].lower()
    ocr = False
    if ext in ('.xlsx', '.xlsm'):
        sheets = rows_from_xlsx(path)
    elif ext == '.xls':
        sheets = rows_from_xls(path)
    elif ext == '.pdf':
        sheets, ocr = rows_from_pdf(path)
    elif ext in ('.png', '.jpg', '.jpeg'):
        sheets, ocr = rows_from_text(ocr_file(path)), True
    else:
        return {}, {}, 'unsupported', False
    best = ({}, {}, None)
    for title, rows in sheets:
        if '견본일지' in title or '체크항목' in title or '공정도' in title or '의뢰서' in title:
            continue
        r, s = extract_from_rows(rows)
        if score(r) > score(best[0]):
            best = (r, s, title)
    return best[0], best[1], best[2], ocr


def rows_from_xls(path):
    import xlrd
    wb = xlrd.open_workbook(path)
    out = []
    for sh in wb.sheets():
        rows = []
        for r in range(sh.nrows):
            toks = [v for v in sh.row_values(r) if str(v).strip() != '']
            if toks:
                rows.append(toks)
        out.append((sh.name, rows))
    return out


# ---------- 파일명 → 품목번호 ----------
def file_numbers(fname):
    """'82~5.' → [82..85], '71. 72.' → [71,72], '57~8.' → [57,58], '116~117.' → [116,117]"""
    b = os.path.basename(fname)
    m = re.match(r'^\s*(\d+)\s*~\s*(\d+)', b)
    if m:
        a, e = m.group(1), m.group(2)
        if len(e) < len(a):
            e = a[:len(a) - len(e)] + e
        return list(range(int(a), int(e) + 1))
    m = re.match(r'^\s*(\d+)\s*\.\s*(\d+)\s*\.', b)
    if m and int(m.group(2)) == int(m.group(1)) + 1:
        return [int(m.group(1)), int(m.group(2))]
    m = re.match(r'^\s*(\d+)', b)
    return [int(m.group(1))] if m else []
