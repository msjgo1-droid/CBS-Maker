"""HOKA CBS Builder — Streamlit app (KO / EN / VI).
One click per batch: formatted internal CBS + buyer CBS for every style uploaded."""
import datetime as dt
import io
import zipfile

import pandas as pd
import streamlit as st

from engine import process_style, group_files, clean_name

st.set_page_config(page_title='HOKA CBS Builder', page_icon='🧾', layout='wide')

LANGS = {'한국어': 'KO', 'English': 'EN', 'Tiếng Việt': 'VI'}
lang = LANGS[st.sidebar.radio('Language / 언어 / Ngôn ngữ', list(LANGS))]

T = {
    'title': ('HOKA CBS 빌더', 'HOKA CBS Builder', 'Công cụ CBS HOKA'),
    'intro': ('파일을 한꺼번에 올리면 스타일번호(7자리)로 자동으로 묶어서, 스타일마다 '
              '**정돈된 Internal CBS + 바이어 CBS** 2개 파일을 만들어요.',
              'Upload everything at once — files are grouped by 7-digit style number and each style gets '
              '**a formatted internal CBS + a buyer CBS**.',
              'Tải tất cả file cùng lúc — hệ thống tự nhóm theo mã style (7 số) và mỗi style tạo ra '
              '**CBS nội bộ đã chuẩn hoá + CBS cho buyer**.'),
    'up_int': ('① Internal CBS (필수, 여러 개 가능)', '① Internal CBS (required, multiple)',
               '① CBS nội bộ (bắt buộc, nhiều file)'),
    'up_base': ('② 바이어 공란 template', '② Buyer blank template', '② Template trống của buyer'),
    'up_prev': ('③ 이전 제출본 (CO 스타일)', '③ Previous submission (CO styles)', '③ Bản đã nộp trước (style CO)'),
    'up_mnmk': ('④ MNMK (YY REQUEST)', '④ MNMK (YY REQUEST)', '④ MNMK (YY REQUEST)'),
    'styles': ('스타일별 파일 묶음', 'Files per style', 'File theo từng style'),
    'run': ('▶ 실행 (Internal 정리 + 바이어 CBS 생성)', '▶ Run (format internal + build buyer CBS)',
            '▶ Chạy (chuẩn hoá CBS nội bộ + tạo CBS buyer)'),
    'crop': ('MNMK 사진에서 마커 부분만 자르기', 'Crop MNMK picture to the marker', 'Cắt ảnh MNMK chỉ lấy sơ đồ'),
    'loose': ('스타일번호를 못 찾았거나 중복된 파일', 'Files without a style number / duplicates',
              'File không có mã style / bị trùng'),
    'no_tpl': ('바이어 template 또는 이전 제출본이 없어서 건너뜀', 'skipped — no buyer template or previous submission',
               'bỏ qua — không có template buyer hoặc bản đã nộp'),
    'no_int': ('Internal CBS가 없어서 건너뜀', 'skipped — no internal CBS', 'bỏ qua — không có CBS nội bộ'),
    'summary': ('결과 요약', 'Summary', 'Tóm tắt'),
    'dl_all': ('⬇️ 전체 다운로드 (zip)', '⬇️ Download all (zip)', '⬇️ Tải tất cả (zip)'),
    'dl_buyer': ('⬇️ 바이어 CBS', '⬇️ Buyer CBS', '⬇️ CBS buyer'),
    'dl_int': ('⬇️ 정돈된 Internal', '⬇️ Formatted internal', '⬇️ CBS nội bộ đã chuẩn hoá'),
    'blocks': ('Internal 블록 (시트별)', 'Internal block per sheet', 'Block nội bộ cho từng sheet'),
    'sketch': ('바이어 CBS 상단 스케치', 'Header sketch', 'Hình phác thảo ở đầu trang'),
    'match': ('항목 매칭 (수정 가능)', 'Item matching (editable)', 'Ghép hạng mục (có thể sửa)'),
    'apply': ('이 스타일 다시 만들기', 'Re-build this style', 'Tạo lại style này'),
    'panels': ('원단(마커) 시트', 'Fabric (marker) sheets', 'Sheet vải (sơ đồ)'),
    'changes': ('변경 내역', 'Changes', 'Thay đổi'),
    'fname': ('파일명', 'File name', 'Tên file'),
    'auto': ('(자동)', '(auto)', '(tự động)'),
    'skip': ('(빼기)', '(skip)', '(bỏ qua)'),
    'legend': ('초록 L = 같은 색 패널이 가로(길이 방향)로 모임 · 빨강 W = 세로(폭 방향)로 쌓임 · 주황 M = 섞임',
               'Green L = same-colour panels run along the length · Red W = stacked across the width · Orange M = mixed',
               'Xanh L = cùng màu nằm theo chiều dài · Đỏ W = xếp theo chiều khổ · Cam M = lẫn lộn'),
    'self_bad': ('🚨 Self(body) MNMK가 1pc / 1way(lengthwise)가 아닌 스타일',
                 '🚨 Styles whose Self (body) marker is NOT 1pc / 1way (lengthwise)',
                 '🚨 Style có sơ đồ Self (thân) KHÔNG phải 1pc / 1way (theo chiều dọc)'),
    'ok': ('문제 없음', 'OK', 'OK'),
    'col_style': ('스타일', 'Style', 'Style'), 'col_mode': ('방식', 'Mode', 'Chế độ'),
    'col_fob': ('바이어 FOB', 'Buyer FOB', 'FOB buyer'), 'col_ifob': ('Internal FOB', 'Internal FOB', 'FOB nội bộ'),
    'col_self': ('Self 1pc/1way', 'Self 1pc/1way', 'Self 1pc/1way'), 'col_warn': ('경고', 'Warnings', 'Cảnh báo'),
}
LI = ['KO', 'EN', 'VI'].index(lang)
t = lambda k: T[k][LI]
VERD = {'lengthwise': ('lengthwise', 'lengthwise', 'chiều dọc'), 'widthwise': ('widthwise (세로 배치)', 'widthwise', 'chiều ngang'),
        'mixed': ('섞여 있음 (random)', 'mixed / random', 'lẫn lộn'), 'unknown': ('판단 불가 (사진 없음)', 'cannot judge (no picture)', 'không xác định (không có ảnh)')}

WARN = {
    'not_1pc': ('Self(body) 마커가 1pc/1way lengthwise가 아님 — [{sheet}] 판정: {verdict} (lengthwise 비율 {share}%)',
                'Self (body) marker is not 1pc/1way lengthwise — [{sheet}] result: {verdict} (lengthwise share {share}%)',
                'Sơ đồ Self (thân) không phải 1pc/1way theo chiều dọc — [{sheet}] kết quả: {verdict} (tỉ lệ chiều dọc {share}%)'),
    'hidden': ('{src} 숨김 시트: {names}', '{src} hidden sheets: {names}', 'Sheet ẩn trong {src}: {names}'),
    'wmf': ('{src}에 WMF/EMF 이미지가 있어 결과 파일에서 빠짐 — 직접 다시 넣어주세요',
            '{src} has WMF/EMF pictures — they are dropped in the output, re-insert manually',
            '{src} có ảnh WMF/EMF — sẽ bị mất trong file kết quả, hãy chèn lại thủ công'),
    'block': ('Internal 블록 #{block}: {text}', 'Internal block #{block}: {text}', 'Block nội bộ #{block}: {text}'),
    'size_block': ('[{sheet}] MNMK에서 "{want}" 사이즈 블록을 못 찾아 마지막 블록 사용',
                   '[{sheet}] MNMK size block "{want}" not found — last block used',
                   '[{sheet}] Không tìm thấy block size "{want}" trong MNMK — dùng block cuối'),
    'panel_nomatch': ('[{sheet}] 원단 시트와 맞는 원단 행이 없어 마커 길이 재계산 안 함',
                      '[{sheet}] no fabric row matched — marker length not recalculated',
                      '[{sheet}] không khớp dòng vải — không tính lại chiều dài sơ đồ'),
    'no_match': ('[{sheet}] {row}행 "{comp}": Internal에 해당 항목 없음 (그대로 둠)',
                 '[{sheet}] row {row} "{comp}": no internal item (left as is)',
                 '[{sheet}] dòng {row} "{comp}": không có trong CBS nội bộ (giữ nguyên)'),
    'split': ('[{sheet}] {row}행 "{comp}": 요척 나눠 넣음 ({src} 비율)', '[{sheet}] row {row} "{comp}": usage split by {src}',
              '[{sheet}] dòng {row} "{comp}": chia định mức theo {src}'),
    'shared': ('[{sheet}] {row}행 "{comp}": 같은 원단({match})이 위 행에 이미 들어감 — Internal에 (a+b) 나눔이 없고 이전 CBS도 안 나눠서 0으로 둠',
               '[{sheet}] row {row} "{comp}": same fabric ({match}) already used on an earlier row — Net Usage set to 0, split manually if needed',
               '[{sheet}] dòng {row} "{comp}": cùng vải ({match}) đã dùng ở dòng trên — Net Usage = 0, tự chia nếu cần'),
    'row_added': ('[{sheet}] {row}행 추가: {comp} ({sec})', '[{sheet}] row {row} added: {comp} ({sec})',
                  '[{sheet}] thêm dòng {row}: {comp} ({sec})'),
    'not_placed': ('[{sheet}] 넣지 못한 Internal 항목: {items}', '[{sheet}] internal items not placed: {items}',
                   '[{sheet}] hạng mục nội bộ chưa đưa vào: {items}'),
    'no_sketch': ('Internal에서 가먼트 스케치를 못 찾음', 'No garment sketch found in internal',
                  'Không tìm thấy hình phác thảo trong CBS nội bộ'),
    'sketch_conv': ('스케치 이미지({ext}) 변환 실패', 'Sketch picture ({ext}) could not be converted',
                    'Không chuyển đổi được ảnh phác thảo ({ext})'),
}


def wtxt(w):
    tpl = WARN.get(w['code'], ('{code}',) * 3)[['KO', 'EN', 'VI'].index(lang)]
    try:
        return tpl.format(**w)
    except Exception:                                    # noqa
        return str(w)


st.title(t('title'))
st.markdown(t('intro'))

c1, c2 = st.columns(2)
f_int = c1.file_uploader(t('up_int'), type=['xlsx'], accept_multiple_files=True, key='int')
f_base = c1.file_uploader(t('up_base'), type=['xlsx'], accept_multiple_files=True, key='base')
f_prev = c2.file_uploader(t('up_prev'), type=['xlsx'], accept_multiple_files=True, key='prev')
f_mnmk = c2.file_uploader(t('up_mnmk'), type=['xlsx'], accept_multiple_files=True, key='mnmk')
crop = st.sidebar.checkbox(t('crop'), value=True)

pack = lambda fs: [(f.name, f.getvalue()) for f in (fs or [])]
groups, loose = group_files(pack(f_int), pack(f_base), pack(f_prev), pack(f_mnmk))

if groups:
    st.subheader(t('styles'))
    rows = []
    for stl, g in sorted(groups.items()):
        mode = ('CO + MNMK' if 'prev' in g and 'mnmk' in g else 'CO' if 'prev' in g else
                'NEW + MNMK' if 'mnmk' in g else 'TEMPLATE')
        rows.append({t('col_style'): stl, t('col_mode'): mode,
                     'Internal': clean_name(g['internal'][0]) if 'internal' in g else '❌',
                     'Template': clean_name(g['base'][0]) if 'base' in g else '',
                     'Previous': clean_name(g['prev'][0]) if 'prev' in g else '',
                     'MNMK': clean_name(g['mnmk'][0]) if 'mnmk' in g else ''})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
if loose:
    st.warning(t('loose') + ': ' + ', '.join(f'{n} ({r})' for r, n in loose))

res = st.session_state.setdefault('results', {})
opts = st.session_state.setdefault('opts', {})


def run_one(stl, g):
    if 'internal' not in g:
        res[stl] = dict(error=t('no_int'))
        return
    if not ('base' in g or 'prev' in g):
        res[stl] = dict(error=t('no_tpl'))
        return
    o = opts.get(stl, {})
    try:
        data, rep = process_style(g, crop=crop, block_map=o.get('block_map'), manual=o.get('manual'),
                                  sketch=o.get('sketch', 0))
        res[stl] = dict(data=data, rep=rep, files=g)
    except Exception as e:                               # noqa
        res[stl] = dict(error=f'{type(e).__name__}: {e}')


if groups and st.button(t('run'), type='primary', use_container_width=True):
    res.clear()
    bar = st.progress(0.0)
    for i, (stl, g) in enumerate(sorted(groups.items())):
        bar.progress(i / len(groups), text=stl)
        run_one(stl, g)
    bar.progress(1.0, text='done')

if res:
    # ---------- 1pc / 1way alert first ----------
    bad = [(s, r['rep']['self_check']) for s, r in res.items() if 'rep' in r
           and r['rep']['self_check']['verdict'] != 'lengthwise']
    if bad:
        st.error(t('self_bad') + '\n\n' + '\n'.join(
            f"- **{s}** — [{c['sheet']}] {VERD[c['verdict']][LI]} ({int((c['share'] or 0) * 100)}%)" for s, c in bad))
        cols = st.columns(min(len(bad), 3))
        for i, (s, c) in enumerate(bad):
            if c['annotated']:
                cols[i % len(cols)].image(c['annotated'], caption=f"{s} · {c['sheet']}", use_container_width=True)
        st.caption(t('legend'))

    st.subheader(t('summary'))
    srows = []
    for s, r in sorted(res.items()):
        if 'error' in r:
            srows.append({t('col_style'): s, t('col_warn'): '❌ ' + r['error']})
            continue
        rep = r['rep']
        sc = rep['self_check']
        srows.append({t('col_style'): s,
                      t('col_fob'): ' / '.join(f"{k}: ${v['fob']:.2f}" for k, v in rep['sheets'].items()),
                      t('col_ifob'): ' / '.join(f"${(v['internal_fob'] or 0):.2f}" for v in rep['sheets'].values()),
                      t('col_self'): '✅' if sc['verdict'] == 'lengthwise' else '🚨 ' + VERD[sc['verdict']][LI],
                      t('col_warn'): len(rep['warnings']), t('fname'): rep['filename']})
    st.dataframe(pd.DataFrame(srows), hide_index=True, use_container_width=True)

    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, 'w', zipfile.ZIP_DEFLATED) as z:
        for s, r in res.items():
            if 'rep' in r:
                z.writestr(f"{s}/{r['rep']['filename']}", r['data'])
                n, d = r['rep']['internal_formatted']
                z.writestr(f'{s}/{n}', d)
    st.download_button(t('dl_all'), zbuf.getvalue(), file_name=f'HOKA_CBS_{dt.date.today():%m%d%Y}.zip',
                       type='primary', use_container_width=True)

    # ---------- per style ----------
    for s, r in sorted(res.items()):
        if 'rep' not in r:
            continue
        rep = r['rep']
        icon = '🚨' if rep['self_check']['verdict'] != 'lengthwise' else ('⚠️' if rep['warnings'] else '✅')
        with st.expander(f"{icon} {s} — {rep['filename']}"):
            for w in rep['warnings']:
                (st.error if w['code'] == 'not_1pc' else st.warning)(wtxt(w))
            if rep['self_check']['annotated']:
                st.image(rep['self_check']['annotated'], width=700,
                         caption=f"Self · {rep['self_check']['sheet']} · {VERD[rep['self_check']['verdict']][LI]}")
                st.caption(t('legend'))
            fname = st.text_input(t('fname'), rep['filename'], key=f'fn_{s}')
            d1, d2 = st.columns(2)
            d1.download_button(t('dl_buyer'), r['data'], file_name=fname, key=f'db_{s}')
            n, d = rep['internal_formatted']
            d2.download_button(t('dl_int'), d, file_name=n, key=f'di_{s}')

            o = opts.setdefault(s, {})
            # blocks
            info = rep['internal']
            st.markdown(f"**{t('blocks')}** · tab `{info['tab']}` · gender `{info['gender']}` · fit `{info['fit']}`")
            bopts = [b['idx'] for b in info['blocks']]
            bm = {}
            cols = st.columns(max(len(rep['sheets']), 1))
            for i, sh in enumerate(rep['sheets']):
                cur = rep['sheets'][sh]['block']
                bm[sh] = cols[i].selectbox(sh, bopts, index=bopts.index(cur), key=f'bl_{s}_{sh}',
                                           format_func=lambda k, info=info: f"#{k} {info['blocks'][k]['stage_hdr'][:25]} "
                                                                 f"${(info['blocks'][k]['fob'] or 0):.2f}"
                                                                 f"{' ✅' if info['blocks'][k]['submitted'] else ''}")
            # sketch
            if rep.get('sketches'):
                st.markdown(f"**{t('sketch')}**")
                sk = st.radio(t('sketch'), list(range(len(rep['sketches']))), index=o.get('sketch', 0),
                              horizontal=True, key=f'sk_{s}', label_visibility='collapsed',
                              format_func=lambda k, rep=rep: f"#{k} r{rep['sketches'][k]['row']} {rep['sketches'][k]['name']}")
                if rep.get('sketch_png'):
                    st.image(rep['sketch_png'], width=260)
            else:
                sk = 0
            # matching
            st.markdown(f"**{t('match')}**")
            man = {}
            for sh, sr in rep['sheets'].items():
                options = [t('auto'), t('skip')] + [f'{i}: {x}' for i, x in enumerate(sr['lines'])]
                df = pd.DataFrame([dict(row=l['row'], component=l['component'], method=l['how'],
                                        internal=f"{l['line']}: {sr['lines'][l['line']]}" if 'line' in l else '',
                                        change=t('auto')) for l in sr['log']])
                st.caption(sh)
                ed = st.data_editor(df, key=f'ed_{s}_{sh}', hide_index=True, use_container_width=True,
                                    disabled=['row', 'component', 'method', 'internal'],
                                    column_config={'change': st.column_config.SelectboxColumn(options=options)})
                m = {}
                for _, x in ed.iterrows():
                    if x['change'] == t('skip'):
                        m[int(x['row'])] = -1
                    elif x['change'] not in (t('auto'), None, ''):
                        m[int(x['row'])] = int(str(x['change']).split(':')[0])
                if m:
                    man[sh] = m
            if st.button(t('apply'), key=f'ap_{s}'):
                o.update(block_map=bm, sketch=sk, manual=man or o.get('manual'))
                run_one(s, r['files'])
                st.rerun()
            # panels
            if rep['panels']:
                st.markdown(f"**{t('panels')}**")
                st.dataframe(pd.DataFrame([dict(sheet=p['panel'], component=p['component'], yy=p['net'],
                                                aux=', '.join(p['aux']), marker=f"{p['marker'][0]} Y {p['marker'][1]} IN"
                                                if p['marker'] else '', check=p['check']) for p in rep['panels']]),
                             hide_index=True, use_container_width=True)
            if st.toggle(t('changes'), key=f'ch_{s}'):
                st.dataframe(pd.DataFrame([dict(c) for v in rep['sheets'].values() for c in v['changes']]),
                             hide_index=True, use_container_width=True)
