"""HOKA CBS Builder — Streamlit app (KO / EN / VI)."""
import datetime as dt
import io

import openpyxl
import pandas as pd
import streamlit as st

from cbs_fill import is_main_sheet
from cbs_format import normalize_bytes
from engine import generate
from internal import read_internal
from panels import read_mnmk, pick_block, crop_marker

st.set_page_config(page_title='HOKA CBS Builder', page_icon='🧾', layout='wide')

T = {
    'title': {'KO': 'HOKA CBS 빌더', 'EN': 'HOKA CBS Builder', 'VI': 'HOKA CBS Builder'},
    'tab_build': {'KO': '① 바이어 CBS 만들기', 'EN': '① Build buyer CBS', 'VI': '① Tạo CBS cho buyer'},
    'tab_fmt': {'KO': '② Internal 서식 정리', 'EN': '② Format internal CBS', 'VI': '② Định dạng CBS nội bộ'},
    'up_int': {'KO': 'Internal CBS (필수)', 'EN': 'Internal CBS (required)', 'VI': 'CBS nội bộ (bắt buộc)'},
    'up_base': {'KO': '바이어 공란 template', 'EN': 'Buyer blank template', 'VI': 'Template trống của buyer'},
    'up_prev': {'KO': '이전 제출본 (CO 스타일)', 'EN': 'Previous submission (CO style)', 'VI': 'Bản đã nộp trước (style CO)'},
    'up_mnmk': {'KO': 'MNMK (YY REQUEST) 파일', 'EN': 'MNMK (YY REQUEST) file', 'VI': 'File MNMK (YY REQUEST)'},
    'need_tpl': {'KO': '바이어 template 또는 이전 제출본 중 하나는 꼭 올려주세요.',
                 'EN': 'Upload a buyer template or a previous submission.',
                 'VI': 'Hãy tải lên template của buyer hoặc bản đã nộp trước.'},
    'mode': {'KO': '처리 방식', 'EN': 'Mode', 'VI': 'Chế độ'},
    'analysis': {'KO': 'Internal 분석', 'EN': 'Internal analysis', 'VI': 'Phân tích CBS nội bộ'},
    'blocks': {'KO': '블록 선택 (시트별)', 'EN': 'Block per sheet', 'VI': 'Chọn block cho từng sheet'},
    'match': {'KO': '항목 매칭 (수정 가능)', 'EN': 'Item matching (editable)', 'VI': 'Ghép hạng mục (có thể sửa)'},
    'apply': {'KO': '매칭 적용해서 다시 만들기', 'EN': 'Re-build with this matching', 'VI': 'Tạo lại với ghép này'},
    'panels': {'KO': '원단(마커) 시트', 'EN': 'Fabric (marker) sheets', 'VI': 'Sheet vải (sơ đồ)'},
    'crop': {'KO': 'MNMK 사진에서 마커 부분만 자르기', 'EN': 'Crop MNMK image to the marker',
             'VI': 'Cắt ảnh MNMK chỉ lấy phần sơ đồ'},
    'fname': {'KO': '파일명', 'EN': 'File name', 'VI': 'Tên file'},
    'download': {'KO': '바이어 CBS 다운로드', 'EN': 'Download buyer CBS', 'VI': 'Tải CBS cho buyer'},
    'changes': {'KO': '변경 내역', 'EN': 'Changes', 'VI': 'Thay đổi'},
    'warn': {'KO': '확인 필요', 'EN': 'Check', 'VI': 'Cần kiểm tra'},
    'fob': {'KO': 'FOB 검증', 'EN': 'FOB check', 'VI': 'Kiểm tra FOB'},
    'fmt_help': {'KO': '1193712 기준 서식: 단가·YY 소수점 최대 4자리, FOB·마진·Net·OH/Profit/Finance 2자리. 값·수식은 그대로.',
                 'EN': 'Standard format (1193712): prices/YY up to 4 decimals; FOB, margin, Net, OH/Profit/Finance 2 decimals. Values untouched.',
                 'VI': 'Định dạng chuẩn (1193712): đơn giá/YY tối đa 4 số lẻ; FOB, margin, Net, OH/Profit/Finance 2 số lẻ. Giữ nguyên giá trị.'},
    'all_tabs': {'KO': '모든 탭 적용 (기본: 맨 앞 탭만)', 'EN': 'All tabs (default: first tab only)',
                 'VI': 'Tất cả các tab (mặc định: tab đầu)'},
    'skip': {'KO': '(빼기)', 'EN': '(skip)', 'VI': '(bỏ qua)'},
    'auto': {'KO': '(자동)', 'EN': '(auto)', 'VI': '(tự động)'},
}
lang = st.sidebar.radio('Language', ['KO', 'EN', 'VI'], horizontal=True)
t = lambda k: T[k][lang]

st.title(t('title'))
tab1, tab2 = st.tabs([t('tab_build'), t('tab_fmt')])


def up_bytes(f):
    return f.getvalue() if f else None


# ======================= TAB 1 =======================
with tab1:
    c1, c2 = st.columns(2)
    f_int = c1.file_uploader(t('up_int'), type=['xlsx'], key='int')
    f_base = c1.file_uploader(t('up_base'), type=['xlsx'], key='base')
    f_prev = c2.file_uploader(t('up_prev'), type=['xlsx'], key='prev')
    f_mnmk = c2.file_uploader(t('up_mnmk'), type=['xlsx'], key='mnmk')

    if f_int and not (f_base or f_prev):
        st.info(t('need_tpl'))
    if f_int and (f_base or f_prev):
        mode = ('CO + MNMK' if f_prev and f_mnmk else 'CO' if f_prev else
                'NEW / MNMK' if f_mnmk else 'TEMPLATE ONLY')
        st.caption(f"{t('mode')}: **{mode}**")

        info, _, _ = read_internal(io.BytesIO(up_bytes(f_int)), f_int.name)
        with st.expander(t('analysis'), expanded=True):
            st.write(f"Tab **{info['tab']}** · style **{info['style']}** · gender **{info['gender']}** · "
                     f"fit stage **{info['fit']}** · buy stage **{info['tab_info']['buy']}**")
            df = pd.DataFrame([dict(block=b['idx'], rows=f"{b['r0']}-{b['r1']}", header=b['stage_hdr'],
                                    FOB=round(b['fob'], 2) if b['fob'] else None,
                                    note=', '.join(f"{d} ${p}" for d, p in b['notes']),
                                    submitted='✅' if b['submitted'] else '—',
                                    heather='Heather' if b['heather'] else '',
                                    warning='; '.join(b['warn'])) for b in info['blocks']])
            st.dataframe(df, hide_index=True, use_container_width=True)
            if info['hidden']:
                st.warning('Internal hidden sheets: ' + ', '.join(info['hidden']))

        # main sheets that will exist in output
        ref = openpyxl.load_workbook(io.BytesIO(up_bytes(f_prev or f_base)))
        titles = [s.title for s in ref.worksheets if is_main_sheet(s)]
        st.subheader(t('blocks'))
        bopts = [b['idx'] for b in info['blocks']]
        from engine import default_block_map
        dmap = default_block_map(info, titles)
        cols = st.columns(max(len(titles), 1))
        bmap = {}
        for i, tt in enumerate(titles):
            bmap[tt] = cols[i].selectbox(tt, bopts, index=bopts.index(dmap[tt]),
                                         format_func=lambda k: f"#{k} {info['blocks'][k]['stage_hdr']} "
                                                               f"${round(info['blocks'][k]['fob'] or 0, 2)}",
                                         key=f'b_{tt}')
        crop = st.checkbox(t('crop'), value=True) if f_mnmk else True

        manual = st.session_state.get('manual', {})
        kw = dict(internal=up_bytes(f_int), internal_name=f_int.name,
                  base=up_bytes(f_base), base_name=f_base.name if f_base else '',
                  prev=up_bytes(f_prev), prev_name=f_prev.name if f_prev else '',
                  mnmk=up_bytes(f_mnmk), block_map=bmap, manual=manual, crop=crop)
        try:
            data, rep = generate(**kw)
        except Exception as e:                                  # noqa
            st.error(f'{type(e).__name__}: {e}')
            st.stop()

        # ---- warnings + FOB ----
        warns = list(rep['warnings'])
        for s, r in rep['sheets'].items():
            if r['unplaced']:
                warns.append(f'[{s}] not placed: ' + ', '.join(r['unplaced']))
            for a in r['added']:
                warns.append(f"[{s}] row added r{a['row']}: {a['component']} ({a['section']})")
            for l in r['log']:
                if l['how'] == 'NO MATCH':
                    warns.append(f"[{s}] r{l['row']} {l['component']}: no internal item")
        if warns:
            st.subheader('⚠️ ' + t('warn'))
            for w in warns:
                st.warning(w)
        st.subheader(t('fob'))
        st.dataframe(pd.DataFrame([dict(sheet=s, block=r['block'], buyer_FOB=round(r['fob'], 4),
                                         internal_FOB=round(r['internal_fob'] or 0, 4),
                                         diff=round(r['fob'] - (r['internal_fob'] or 0), 4),
                                         changes=len(r['changes'])) for s, r in rep['sheets'].items()]),
                     hide_index=True, use_container_width=True)

        # ---- matching editor ----
        st.subheader(t('match'))
        new_manual = {}
        for s, r in rep['sheets'].items():
            opts = [t('auto'), t('skip')] + [f'{i}: {x}' for i, x in enumerate(r['lines'])]
            rows = []
            for l in r['log']:
                cur = f"{l['line']}: {r['lines'][l['line']]}" if 'line' in l else ''
                rows.append(dict(row=l['row'], component=l['component'], method=l['how'],
                                 internal=cur, change=t('auto')))
            ed = st.data_editor(pd.DataFrame(rows), key=f'ed_{s}', hide_index=True, use_container_width=True,
                                disabled=['row', 'component', 'method', 'internal'],
                                column_config={'change': st.column_config.SelectboxColumn(options=opts)})
            m = {}
            for _, x in ed.iterrows():
                if x['change'] == t('skip'):
                    m[int(x['row'])] = -1
                elif x['change'] not in (t('auto'), None, ''):
                    m[int(x['row'])] = int(str(x['change']).split(':')[0])
            if m:
                new_manual[s] = m
        if st.button(t('apply')):
            st.session_state['manual'] = new_manual
            st.rerun()

        # ---- panels ----
        st.subheader(t('panels'))
        if rep['panels']:
            st.dataframe(pd.DataFrame([dict(sheet=p['panel'], component=p['component'], yy=p['net'],
                                            aux=', '.join(p['aux']), aux_yy=p['aux_yy'],
                                            marker=f"{p['marker'][0]} Y {p['marker'][1]} IN" if p['marker'] else '',
                                            check=p['check']) for p in rep['panels']]),
                         hide_index=True, use_container_width=True)
        if f_mnmk:
            mwb = openpyxl.load_workbook(io.BytesIO(up_bytes(f_mnmk)), data_only=True)
            ms, _ = read_mnmk(mwb)
            pc = st.columns(min(3, max(len(ms), 1)))
            for i, m in enumerate(ms):
                b, _ = pick_block(m['blocks'], info['gender'])
                if b and b['image'] is not None:
                    img = b['image']._data()
                    if crop:
                        img = crop_marker(img)[0]
                    pc[i % len(pc)].image(img, caption=f"{m['title']} · {b['marker_y']}Y {b['marker_in']}in · "
                                                       f"eff {b['eff']}", use_container_width=True)

        with st.expander(t('changes')):
            ch = [dict(c) for r in rep['sheets'].values() for c in r['changes']]
            st.dataframe(pd.DataFrame(ch), hide_index=True, use_container_width=True)

        fname = st.text_input(t('fname'), rep['filename'])
        st.download_button('⬇️ ' + t('download'), data, file_name=fname, type='primary',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ======================= TAB 2 =======================
with tab2:
    st.caption(t('fmt_help'))
    files = st.file_uploader(t('up_int'), type=['xlsx'], accept_multiple_files=True, key='fmt')
    all_tabs = st.checkbox(t('all_tabs'))
    for f in files or []:
        data_out, log = normalize_bytes(f.getvalue(), all_sheets=all_tabs)
        skipped = [l for l in log if '!' in l]
        st.write(f"**{f.name}** — {sum('ok' in l for l in log)} blocks" + (f" · ⚠️ {len(skipped)}" if skipped else ''))
        for l in skipped:
            st.warning(l)
        st.download_button('⬇️ ' + f.name, data_out, file_name=f.name, key=f'dl_{f.name}')
