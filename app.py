"""HOKA CBS Builder — Streamlit app (KO / EN / VI).
One click per batch: formatted internal CBS + buyer CBS for every style uploaded."""
import datetime as dt
import io
import zipfile

import pandas as pd
import streamlit as st

from engine import process_style, group_files, clean_name
import store as fstore
import cif as fcif

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
    'tab_make': ('🧾 CBS 만들기', '🧾 Build CBS', '🧾 Tạo CBS'),
    'tab_upd': ('🗂 Internal CBS 업데이트', '🗂 Internal CBS update', '🗂 Cập nhật CBS nội bộ'),
    'upd_intro': ('Internal CBS만 올리세요. 맨 왼쪽(최신) 탭을 이전 단계 탭(날짜가 다른 다음 탭)과 비교해서, 단가·요척·LOSS·CIF·CM이 바뀐 곳을 '
                  'REMARK 칸에 **빨간 글씨**로 적어요. 기존 REMARK 내용은 그대로 두고 뒤에 이어 적어요. 가격(내부/오픈)은 절대 바꾸지 않아요. '
                  '숫자 서식·소계 정리와 CIF 확인도 같이 해요.',
                  'Upload internal CBS only. The leftmost (latest) tab is compared with the previous stage tab (next tab with another date); '
                  'what changed (price, yy, loss, CIF, CM) is written in the REMARK cell in **red**, after what is already there. '
                  'Prices (internal/open) are never changed. Number formats, subtotals and CIF check are done too.',
                  'Chỉ tải CBS nội bộ. Tab ngoài cùng bên trái (mới nhất) được so với tab giai đoạn trước; thay đổi (giá, định mức, '
                  'hao hụt, CIF, CM) được ghi vào ô REMARK bằng **chữ đỏ**, sau nội dung sẵn có. Giá không bao giờ bị sửa.'),
    'upd_open0': ('올린 파일에서 이미 Open YY가 0인 항목 (ACTUAL LOSS 칸이 비어 있어서 Open 금액이 빠져 있어요 — 앱은 값을 바꾸지 않아요)',
                  'Open YY already 0 in the uploaded file (ACTUAL LOSS empty, so the open cost is missing — the app changes no values)',
                  'Open YY đã = 0 trong file tải lên (ACTUAL LOSS trống — app không sửa giá trị)'),
    'upd_run': ('▶ 업데이트 실행', '▶ Update', '▶ Cập nhật'),
    'upd_prev': ('비교한 이전 단계', 'Compared with', 'So với'),
    'upd_none': ('이전 단계 탭(날짜가 다른 탭)이 없어서 비교하지 못했어요', 'No previous-stage tab (another date) to compare with',
                 'Không có tab giai đoạn trước để so sánh'),
    'upd_notes': ('REMARK에 적은 내용', 'Written in REMARK', 'Đã ghi vào REMARK'),
    'upd_gone': ('이전 단계에는 있었는데 지금 없는 항목', 'Items in the previous stage that are gone now', 'Mục có ở giai đoạn trước nhưng nay không còn'),
    'upd_same': ('이전 단계와 바뀐 것이 없어요', 'Nothing changed vs the previous stage', 'Không có thay đổi so với giai đoạn trước'),
    'upd_skip': ('REMARK가 수식이라 적지 않은 곳', 'REMARK is a formula — not written', 'REMARK là công thức — không ghi'),
    'tab_hist': ('📊 FOB 기록', '📊 FOB history', '📊 Lịch sử FOB'),
    'save_hist': ('💾 이 결과를 FOB 기록에 저장', '💾 Save these results to FOB history', '💾 Lưu kết quả vào lịch sử FOB'),
    'saved': ('저장 완료: 새로 {a}건, 갱신 {r}건', 'Saved: {a} new, {r} updated', 'Đã lưu: {a} mới, {r} cập nhật'),
    'store_local': ('⚠️ 저장소가 설정되지 않아 서버 임시 폴더에 저장돼요 (Streamlit Cloud에서는 재시작 시 사라짐). README의 [fob_store] 설정을 해주세요.',
                    '⚠️ No storage configured — saved to a temporary folder (lost when Streamlit Cloud restarts). Set [fob_store] in secrets (see README).',
                    '⚠️ Chưa cấu hình lưu trữ — dữ liệu lưu tạm (mất khi Streamlit Cloud khởi động lại). Cài [fob_store] trong secrets (xem README).'),
    'import': ('예전에 바이어에게 나간 CBS 올려서 기록에 추가', 'Add previously submitted buyer CBS files',
               'Thêm các file CBS đã gửi buyer trước đây'),
    'import_hint': ('파일명/시트에서 시즌·단계·스타일·컬러·FOB를 읽어요. 아래 표에서 고친 다음 저장하세요.',
                    'Season / stage / style / colour / FOB are read from the file name and sheets. Fix them below, then save.',
                    'Mùa / giai đoạn / style / màu / FOB đọc từ tên file và sheet. Sửa bên dưới rồi lưu.'),
    'save': ('💾 저장', '💾 Save', '💾 Lưu'),
    'f_season': ('시즌', 'Season', 'Mùa'), 'f_stage': ('단계', 'Stage', 'Giai đoạn'),
    'f_style': ('스타일 검색', 'Style search', 'Tìm style'),
    'compare': ('단계별 비교 (선택한 단계를 열로)', 'Compare stages (stages as columns)', 'So sánh giai đoạn (mỗi giai đoạn một cột)'),
    'dl_xlsx': ('⬇️ 엑셀로 받기', '⬇️ Download Excel', '⬇️ Tải Excel'),
    'empty': ('저장된 기록이 없어요.', 'No records yet.', 'Chưa có dữ liệu.'),
    'chart': ('차트', 'Chart', 'Biểu đồ'),
    'up_int': ('① Internal CBS (필수, 여러 개 가능)', '① Internal CBS (required, multiple)',
               '① CBS nội bộ (bắt buộc, nhiều file)'),
    'up_base': ('② 바이어 공란 template', '② Buyer blank template', '② Template trống của buyer'),
    'up_prev': ('③ 이전 제출본 (CO 스타일)', '③ Previous submission (CO styles)', '③ Bản đã nộp trước (style CO)'),
    'up_mnmk': ('④ MNMK (YY REQUEST)', '④ MNMK (YY REQUEST)', '④ MNMK (YY REQUEST)'),
    'styles': ('스타일별 파일 묶음', 'Files per style', 'File theo từng style'),
    'cif_bad': ('Internal CIF 확인 필요 — 지금까지 다른 스타일에 같은 품목으로 넣은 CIF와 달라요',
                'Check the internal CIF — different from what other styles used for the same kind of item',
                'Kiểm tra CIF nội bộ — khác với CIF các style khác dùng cho cùng loại hàng'),
    'cif_type': ('품목 종류', 'Item type', 'Loại hàng'),
    'cif_usual': ('다른 스타일 (값×건수)', 'Other styles (value×count)', 'Style khác (giá trị×số lần)'),
    'cif_why': ('이유', 'Why', 'Lý do'),
    'cif_missing': ('CIF 빠짐', 'CIF missing', 'Thiếu CIF'),
    'cif_differs': ('같은 품목과 다름', 'Differs from the same item', 'Khác cùng loại hàng'),
    'cif_high': ('평소보다 높음', 'Higher than usual', 'Cao hơn thường lệ'),
    'cif_low': ('평소보다 낮음', 'Lower than usual', 'Thấp hơn thường lệ'),
    'cif_hist': ('CIF 기록 (Internal CBS)', 'CIF records (internal CBS)', 'Lịch sử CIF (CBS nội bộ)'),
    'cif_hist_hint': ('CBS를 만들 때마다 Internal의 CIF가 자동 저장돼요. 예전 Internal 파일을 여기 올려서 기록에 추가할 수도 있어요.',
                      'Every run saves the internal CIF automatically. You can also add old internal files here.',
                      'Mỗi lần chạy CIF nội bộ được lưu tự động. Có thể thêm file nội bộ cũ ở đây.'),
    'qty_title': ('원단 시트 수량 (MNMK에서 못 읽은 칸 직접 입력)', 'Fabric sheet quantities (type what the MNMK picture could not be read for)',
                  'Số lượng sheet vải (nhập ô không đọc được từ ảnh MNMK)'),
    'qty_hint': ("빈칸(Garment Q'ty 등)을 사진 보고 입력한 뒤 아래 '이 스타일 다시 만들기'를 누르면 마커 길이가 계산돼요",
                 "Fill the blanks (Garment Q'ty etc.) from the picture, then press 'Re-build this style' — the marker length is calculated",
                 "Nhập các ô trống (Garment Q'ty…) theo ảnh rồi bấm 'Tạo lại style này' — chiều dài sơ đồ sẽ được tính"),
    'reading_mnmk': ('MNMK 사진 읽는 중', 'reading MNMK picture', 'đang đọc ảnh MNMK'),
    'working': ('작업 중…', 'working…', 'đang xử lý…'),
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
    'drop': ('📉 이전 단계보다 낮아진 항목', '📉 Items lower than previous stage (tick = keep previous)',
             '📉 Hạng mục thấp hơn giai đoạn trước (tích = giữ giá trị cũ)'),
    'yy_title': ('📉 이전 날짜 탭보다 요척(Internal YY)이 줄어든 항목', '📉 Internal YY lower than the previous-date tab',
                 '📉 Định mức nội bộ (YY) thấp hơn tab ngày trước'),
    'yy_none': ('요척이 줄어든 항목이 없어요.', 'No YY decreases.', 'Không có định mức giảm.'),
    'ok_btn': ('확인', 'OK', 'OK'),
    'pic_title': ('🖼 원단 시트 MNMK 사진 출처', '🖼 Marker picture source per fabric sheet', '🖼 Nguồn ảnh sơ đồ cho từng sheet vải'),
    'pic_hint': ('자동: SELF는 1pc 1way lengthwise인 쪽, 나머지 원단은 Internal 사진. Internal 사진을 쓰면 폭·효율·Q\'ty도 그 사진 값으로 바뀌어요. YY는 항상 Open YY 기준.',
                 'Auto: SELF uses whichever is 1pc/1way lengthwise, other fabrics use the internal picture. Using the internal picture also takes its width / efficiency / Q\'ty. YY always follows the open YY.',
                 'Tự động: SELF chọn ảnh 1pc/1way theo chiều dọc, vải khác dùng ảnh nội bộ. Dùng ảnh nội bộ thì khổ / hiệu suất / SL cũng theo ảnh đó. YY luôn theo YY open.'),
    'pic_prev': ('이전 제출 CBS', 'Previous CBS', 'CBS đã nộp'),
    'basis': ('💲 바이어 CBS에 쓸 Open 가격 기준', '💲 Buyer open price basis', '💲 Chọn cơ sở giá open cho buyer'),
    'basis_hint': ('Internal 파일의 내부(Internal)·오픈(Open) 가격은 바뀌지 않아요. 선택한 기준은 바이어 CBS에만 적용돼요 → 고른 뒤 "이 스타일 다시 만들기".',
                   'The internal file prices never change; the choice only applies to the buyer CBS → then "Re-build this style".',
                   'Giá trong file nội bộ không đổi; lựa chọn chỉ áp dụng cho CBS buyer → sau đó bấm "Tạo lại style này".'),
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
PICSRC = {'auto': ('자동', 'Auto', 'Tự động'), 'internal': ('Internal 사진', 'Internal picture', 'Ảnh nội bộ'),
          'prev': ('이전 제출 CBS 사진', 'Previous CBS picture', 'Ảnh CBS đã nộp')}
BASIS = {'current': ('현재 Open 그대로 (기본)', 'Current internal open (default)', 'Giá open nội bộ hiện tại (mặc định)'),
         'prev_yy': ('Open YY만 이전 단계 값 유지', 'Keep previous open YY only', 'Giữ YY open giai đoạn trước'),
         'prev_all': ('낮아진 Open 항목 모두 이전 값 유지', 'Keep previous open for every lower item', 'Giữ open cũ cho mọi hạng mục giảm'),
         'custom': ('항목별 직접 선택', 'Choose per item', 'Chọn từng hạng mục')}
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
    'price_drop': ('[{sheet}] 바이어 open FOB가 이전 단계({tab}) ${prev} → ${cur}로 낮아짐 — 아래에서 원인 확인 후 이전 값 유지 선택 가능',
                   '[{sheet}] open FOB dropped vs previous stage ({tab}) ${prev} → ${cur} — check the cause below and choose to keep previous values',
                   '[{sheet}] Giá FOB open giảm so với giai đoạn trước ({tab}) ${prev} → ${cur} — xem nguyên nhân bên dưới, có thể giữ giá trị cũ'),
    'no_5pct': ('[{sheet}] MNMK에 shrinkage 5% 마커가 없어 다른 블록 사용', '[{sheet}] no 5% shrinkage marker in MNMK — another block used',
                '[{sheet}] MNMK không có sơ đồ co rút 5% — dùng block khác'),
    'from_internal': ('이전 제출본·MNMK 파일이 없어 Internal 우측 MNMK 사진 {n}장으로 원단 시트를 만들었어요 — 사진에서 읽은 값(폭·효율·Q\'ty)을 확인해 주세요',
                      'No previous submission / MNMK file — {n} fabric sheets built from the MNMK pictures on the right of the internal; please check the values read from the pictures',
                      'Không có bản nộp trước / file MNMK — đã tạo {n} sheet vải từ ảnh MNMK bên phải CBS nội bộ; hãy kiểm tra các giá trị đọc từ ảnh'),
    'sheet_added': ('[{sheet}] 업로드한 바이어 CBS에 없던 원단 시트 — 내부 CBS의 MNMK 사진으로 새로 만들었어요. 값 확인해주세요',
                    '[{sheet}] fabric sheet missing from the uploaded buyer CBS — built from the internal MNMK picture; please check the values',
                    '[{sheet}] sheet vải chưa có trong CBS buyer đã tải lên — đã tạo từ ảnh MNMK của CBS nội bộ; hãy kiểm tra giá trị'),
    'split_part': ('[{sheet}] 같은 원단의 다른 부위 MNMK — 내부 요척 {part} × 오픈 LOSS 비율 = {yy} 로 계산, 메인 시트는 나머지(오픈 요척 − 이 값)',
                   '[{sheet}] second MNMK of the same fabric — internal yy {part} × open loss ratio = {yy}; the main sheet takes the rest (open yy − this)',
                   '[{sheet}] MNMK thứ hai của cùng loại vải — định mức nội bộ {part} × tỉ lệ hao hụt open = {yy}; sheet chính lấy phần còn lại'),
    'old_files': ('GitHub의 앱 파일 버전이 서로 달라요 — zip 안의 모든 .py 파일을 같이 덮어쓰고 Reboot 해주세요',
                  'App files on GitHub are from different versions — overwrite ALL .py files from the zip and reboot',
                  'Các file app trên GitHub khác phiên bản — ghi đè TẤT CẢ file .py trong zip rồi Reboot'),
    'cif_bad': ('Internal CIF {n}곳이 다른 스타일의 같은 품목 CIF와 달라요 — Internal 파일에 빨간색으로 표시하고 옆에 보통 값을 적었어요',
                'Internal CIF differs from the same kind of item in other styles at {n} place(s) — marked red in the internal file with the usual value next to it',
                'CIF nội bộ khác cùng loại hàng ở style khác tại {n} chỗ — đã tô đỏ trong file nội bộ và ghi giá trị thường dùng bên cạnh'),
    'cif_err': ('CIF 확인 중 오류: {err}', 'CIF check failed: {err}', 'Lỗi kiểm tra CIF: {err}'),
    'sum_fixed': ('[{sheet}] 템플릿 합계 수식이 다른 열(A열 등)을 가리키고 있어서 Unit Price 열로 고쳤어요 ({n}곳) — 안 고치면 FOB가 0으로 나와요',
                  '[{sheet}] template SUM formulas pointed at the wrong column (e.g. A) — re-pointed {n} of them to Unit Price (otherwise FOB shows 0)',
                  '[{sheet}] công thức tổng của template trỏ sai cột (vd. cột A) — đã sửa {n} chỗ về cột Unit Price (nếu không FOB = 0)'),
    'mnmk_tab': ('현재 탭에 MNMK 사진이 없어 [{tab}] 탭의 MNMK 사진을 사용했어요 — 마커 길이는 현재 오픈 요척으로 다시 계산했지만, 최신 마커인지 확인해 주세요',
                 'No MNMK picture on the current tab — used the one on tab [{tab}]; marker length re-calculated from the current open yy, please check it is the latest marker',
                 'Tab hiện tại không có ảnh MNMK — đã dùng ảnh ở tab [{tab}]; chiều dài sơ đồ được tính lại theo định mức open hiện tại, hãy kiểm tra có phải sơ đồ mới nhất không'),
    'ocr_check': ('[{sheet}] 사진에서 못 읽었거나 맞지 않는 값: {miss} — 직접 확인/입력 필요',
                  '[{sheet}] could not read / inconsistent: {miss} — check manually',
                  '[{sheet}] không đọc được / không khớp: {miss} — kiểm tra thủ công'),
    'no_panels': ('원단(마커) 시트가 없어요 — 이전 제출본이나 MNMK 파일을 같이 올리거나, Internal 우측에 MNMK 사진을 넣어주세요',
                  'No fabric (marker) sheets — upload a previous submission or MNMK file, or paste MNMK pictures on the right of the internal',
                  'Không có sheet vải — tải bản nộp trước hoặc file MNMK, hoặc dán ảnh MNMK bên phải CBS nội bộ'),
    'recalc': ('Internal 파일에 계산된 값이 없어서(엑셀에서 저장 안 된 파일) 다시 계산해서 읽었어요 — 가능하면 엑셀에서 한 번 열어 저장한 파일을 올려주세요',
               'The internal file had no calculated values (never saved in Excel) — recalculated to read it; ideally open & save it in Excel first',
               'File nội bộ chưa có giá trị tính toán (chưa lưu bằng Excel) — đã tính lại để đọc; nên mở và lưu bằng Excel trước'),
    'basis_prev': ('[{sheet}] 바이어 CBS를 이전 단계({tab}) Open 기준으로 만듦: 바이어 CBS ${buyer} / Internal 파일 Open ${internal} (Internal 파일은 그대로)',
                   '[{sheet}] buyer open built on previous stage ({tab}): buyer ${buyer} / internal open ${internal} (internal file unchanged)',
                   '[{sheet}] Open buyer theo giai đoạn trước ({tab}): buyer ${buyer} / nội bộ ${internal} (file nội bộ giữ nguyên)'),
    'fob_gap': ('[{sheet}] 바이어 CBS FOB ${buyer}가 Internal 파일의 Open FOB ${internal}와 {diff} 차이 (허용 ±$0.01 초과) — 아래 매칭 표 확인',
                '[{sheet}] buyer FOB ${buyer} differs from internal open FOB ${internal} by {diff} (over ±$0.01) — check the matching table',
                '[{sheet}] FOB buyer ${buyer} lệch {diff} so với FOB nội bộ ${internal} (vượt ±$0.01) — kiểm tra bảng ghép'),
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


def money(v):
    return f'${v:.2f}' if isinstance(v, (int, float)) else '$ -'


def wtxt(w):
    tpl = WARN.get(w['code'], ('{code}',) * 3)[['KO', 'EN', 'VI'].index(lang)]
    try:
        return tpl.format(**w)
    except Exception:                                    # noqa
        return str(w)


def _secrets():
    try:
        return dict(st.secrets)
    except Exception:                                        # noqa
        return None


st.title(t('title'))
tab_make, tab_upd, tab_hist = st.tabs([t('tab_make'), t('tab_upd'), t('tab_hist')])

def cif_history(reload=False):
    """CIF records saved by earlier runs (the engine adds the starter set shipped with the app)."""
    if reload or 'cif_hist' not in st.session_state:
        try:
            st.session_state['cif_hist'] = fcif.CifStore(_secrets()).load()
        except Exception:                                    # noqa
            st.session_state['cif_hist'] = pd.DataFrame(columns=fcif.COLS)
    return st.session_state['cif_hist']


with tab_make:
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
            kw = dict(crop=crop, block_map=o.get('block_map'), manual=o.get('manual'),
                      sketch=o.get('sketch', 0), keep=o.get('keep'), pics=o.get('pics'), qty=o.get('qty'),
                      cif_hist=cif_history())
            import inspect as _insp
            import engine as _eng
            ok = set(_insp.signature(_eng.generate).parameters)
            # app.py and engine.py from different versions (only some files updated on GitHub) -> don't crash
            kw = {k: v for k, v in kw.items() if k in ok or k == 'crop' or (k == 'cif_hist' and hasattr(_eng, 'cif_marks'))}
            data, rep = process_style(g, **kw)
            if 'qty' not in ok:
                rep['warnings'].append(dict(code='old_files'))
            res[stl] = dict(data=data, rep=rep, files=g)
        except Exception as e:                               # noqa
            res[stl] = dict(error=f'{type(e).__name__}: {e}')


    def yy_table(d):
        return pd.DataFrame([{'Item': x['item'], 'Detail': x['detail'][:40],
                              'YY ' + d['prev_tab']: x['prev_yy'], 'YY ' + d['cur_tab']: x['cur_yy'],
                              'Δ': x['diff'], 'Δ%': x['pct'], 'open YY (prev→now)': f"{x['prev_open']:g} → {x['cur_open']:g}"}
                             for x in d['rows']])

    @st.dialog(T['yy_title'][LI], width='large')
    def yy_dialog():
        for s_, r_ in sorted(res.items()):
            for sh_, d_ in (r_.get('rep') or {}).get('yy_drops', {}).items():
                st.markdown(f"**{s_}** · [{sh_}] `{d_['prev_tab']}` → `{d_['cur_tab']}`")
                st.dataframe(yy_table(d_), hide_index=True, use_container_width=True)
        if st.button(t('ok_btn'), type='primary'):
            st.rerun()

    if groups and st.button(t('run'), type='primary', use_container_width=True):
        st.session_state['yy_popup'] = True
        res.clear()
        import time as _time
        import internal_mnmk as _im
        bar = st.progress(0.0)
        n_ = len(groups)
        t0 = _time.time()
        for i, (stl, g) in enumerate(sorted(groups.items())):
            def _cb(k, n, i=i, stl=stl):
                bar.progress(min((i + (k + 1) / max(n, 1) * 0.9) / n_, 1.0),
                             text=f"{stl} ({i + 1}/{n_}) · {t('reading_mnmk')} {k + 1}/{n} · {int(_time.time() - t0)}s")
            _im.PROGRESS = _cb
            bar.progress(i / n_, text=f"{stl} ({i + 1}/{n_}) · {t('working')} · {int(_time.time() - t0)}s")
            run_one(stl, g)
        _im.PROGRESS = None
        bar.progress(1.0, text=f"done · {int(_time.time() - t0)}s")
        allrecs = [x for r_ in res.values() for x in (r_.get('rep') or {}).get('cif_recs') or []]
        if allrecs:
            try:
                fcif.CifStore(_secrets()).upsert(allrecs)        # every internal CIF is kept for later checks
                cif_history(reload=True)
            except Exception as e:                           # noqa
                st.caption(f'CIF history not saved: {type(e).__name__}: {e}')

    if res and st.session_state.pop('yy_popup', False) and any(
            (r.get('rep') or {}).get('yy_drops') for r in res.values()):
        yy_dialog()
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
                          '📉': '📉' if rep.get('drops') else '',
                          'YY↓': sum(len(d['rows']) for d in rep.get('yy_drops', {}).values()) or 0,
                          'CIF⚠': len(rep.get('cif_issues') or []),
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
        if st.button(t('save_hist'), use_container_width=True):
            recs = []
            for s_, r_ in res.items():
                if 'rep' in r_:
                    fn = st.session_state.get(f'fn_{s_}') or r_['rep']['filename']
                    recs += fstore.records_from_buyer_file(fn, r_['data'],
                                                           {k: v['fob'] for k, v in r_['rep']['sheets'].items()})
            if recs:
                stor = fstore.Store(_secrets())
                a, rp = stor.upsert(recs)
                st.success(t('saved').format(a=a, r=rp))
                if stor.kind == 'local':
                    st.warning(t('store_local'))

        # ---------- per style ----------
        for s, r in sorted(res.items()):
            if 'rep' not in r:
                continue
            rep = r['rep']
            icon = '🚨' if rep['self_check']['verdict'] != 'lengthwise' else ('⚠️' if rep['warnings'] else '✅')
            with st.expander(f"{icon} {s} — {rep['filename']}"):
                for w in rep['warnings']:
                    (st.error if w['code'] in ('not_1pc', 'fob_gap') else st.info if w['code'] == 'basis_prev'
                 else st.warning)(wtxt(w))
                if rep.get('cif_issues'):
                    st.error(t('cif_bad'))
                    st.dataframe(pd.DataFrame([{t('cif_type'): x['type'], 'Item': x['detail'], 'Price': x.get('price'),
                                                'CIF': x['cif'], t('cif_usual'): x['seen'],
                                                t('cif_why'): t('cif_' + x['why'])} for x in rep['cif_issues']]),
                                 hide_index=True, use_container_width=True)
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
                for sh, d in rep.get('yy_drops', {}).items():
                    st.markdown(f"**{t('yy_title')}** · [{sh}] `{d['prev_tab']}` → `{d['cur_tab']}`")
                    st.dataframe(yy_table(d), hide_index=True, use_container_width=True)
                # marker picture source per fabric sheet
                pics = {}
                if rep.get('pic_choice'):
                    st.markdown(f"**{t('pic_title')}**")
                    st.caption(t('pic_hint'))
                    VER = {'lengthwise': '✅ lengthwise', 'mixed': '⚠️ mixed', 'widthwise': '⚠️ widthwise', 'unknown': '—'}
                    for pc in rep['pic_choice']:
                        c1_, c2_, c3_ = st.columns([1.2, 2, 2])
                        cur = (o.get('pics') or {}).get(pc['sheet'], 'auto')
                        pics[pc['sheet']] = c1_.radio(pc['sheet'], ['auto', 'internal', 'prev'],
                                                      index=['auto', 'internal', 'prev'].index(cur),
                                                      key=f"pc_{s}_{pc['sheet']}",
                                                      format_func=lambda k: PICSRC[k][LI])
                        c1_.caption(f"→ {PICSRC[pc['used']][LI]}")
                        if pc['internal_png']:
                            c2_.image(pc['internal_png'], caption=f"Internal · {VER.get(pc['internal'], '')}",
                                      use_container_width=True)
                        if pc['previous_png']:
                            c3_.image(pc['previous_png'], caption=f"{t('pic_prev')} · {VER.get(pc['previous'], '')}",
                                      use_container_width=True)
                # buyer OPEN price basis (internal file itself is never changed)
                keep = {}
                for sh, dr in rep.get('drops', {}).items():
                    st.markdown(f"**{t('basis')}** · [{sh}] `{dr['prev_tab']}` {money(dr['prev_fob'])} → "
                                f"`{rep['internal']['tab']}` {money(dr['cur_fob'])}")
                    st.caption(t('basis_hint'))
                    cur = (o.get('keep') or {}).get(sh, 'current')
                    opts_ = ['current', 'prev_yy', 'prev_all', 'custom']
                    pick = st.radio(sh, opts_, index=opts_.index(cur if isinstance(cur, str) else 'custom'),
                                    key=f'bs_{s}_{sh}', horizontal=True, label_visibility='collapsed',
                                    format_func=lambda k: BASIS[k][LI])
                    rows = []
                    ck = cur if isinstance(cur, dict) else {}
                    for x in dr['rows']:
                        k = x['line']
                        rows.append(dict(line=str(k), item=x['item'], detail=x['detail'][:40], reason=', '.join(x['why']),
                                         previous=', '.join(f"{a}={v:g}" for a, v in x['prev'].items() if v is not None),
                                         now=', '.join(f"{a}={v:g}" for a, v in x['cur'].items() if v is not None),
                                         delta=x['delta'],
                                         keep_yy='yy' in ck.get(k, []), keep_price='price' in ck.get(k, []),
                                         keep_loss='loss' in ck.get(k, []), keep_cif='cif' in ck.get(k, []),
                                         keep_cm='cm' in ck.get(k, [])))
                    if pick == 'custom':
                        ed = st.data_editor(pd.DataFrame(rows), key=f'dr_{s}_{sh}', hide_index=True,
                                            use_container_width=True,
                                            disabled=['line', 'item', 'detail', 'reason', 'previous', 'now', 'delta'])
                        kk = {}
                        for _, x in ed.iterrows():
                            f = [a for a in ('yy', 'price', 'loss', 'cif', 'cm') if x[f'keep_{a}']]
                            if f:
                                kk['CM' if x['line'] == 'CM' else int(x['line'])] = f
                        keep[sh] = kk or 'current'
                    else:
                        st.dataframe(pd.DataFrame(rows).drop(columns=[c for c in rows[0] if c.startswith('keep_')]),
                                     hide_index=True, use_container_width=True)
                        keep[sh] = pick
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
                # Garment / Panel Q'ty that could not be read from the MNMK picture -> type them here
                qty = dict(o.get('qty') or {})
                pf = rep.get('panel_fields') or []
                if pf:
                    st.markdown(f"**{t('qty_title')}**")
                    st.caption(t('qty_hint'))
                    qdf = pd.DataFrame([{'sheet': x['sheet'], 'Width': str(x['width'] or ''),
                                         'Efficiency': x['eff'], "Garment Q'ty": x['gqty'], "Panel Q'ty": x['pqty']}
                                        for x in pf])
                    qed = st.data_editor(qdf, key=f'qty_{s}', hide_index=True, use_container_width=True,
                                         disabled=['sheet'],
                                         column_config={"Garment Q'ty": st.column_config.NumberColumn(min_value=1, step=1),
                                                        "Panel Q'ty": st.column_config.NumberColumn(min_value=1, step=1)})
                    for (_, a0), (_, a1) in zip(qdf.iterrows(), qed.iterrows()):
                        ch = {}
                        for col, k in (('Width', 'width'), ('Efficiency', 'eff'), ("Garment Q'ty", 'gqty'), ("Panel Q'ty", 'pqty')):
                            v0, v1 = a0[col], a1[col]
                            if pd.isna(v1) if not isinstance(v1, str) else v1 == '':
                                continue
                            if str(v1) != str(v0):
                                ch[k] = int(v1) if k in ('gqty', 'pqty') else (float(v1) if k == 'eff' else str(v1).strip())
                        if ch:
                            qty[a1['sheet']] = {**qty.get(a1['sheet'], {}), **ch}
                if st.button(t('apply'), key=f'ap_{s}'):
                    o.update(block_map=bm, sketch=sk, manual=man or o.get('manual'), keep=keep, pics=pics or o.get('pics'),
                             qty=qty or None)
                    run_one(s, r['files'])
                    st.rerun()
                if rep.get('internal_mnmk'):
                    st.caption('Internal MNMK OCR')
                    st.dataframe(pd.DataFrame(rep['internal_mnmk']), hide_index=True, use_container_width=True)
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


# ======================= INTERNAL CBS UPDATE =======================
with tab_upd:
    import history_update as hupd
    st.markdown(t('upd_intro'))
    ups = st.file_uploader('Internal CBS', type=['xlsx'], accept_multiple_files=True, key='upd_files')
    ures = st.session_state.setdefault('upd_res', {})
    if ups and st.button(t('upd_run'), type='primary', use_container_width=True, key='upd_go'):
        ures.clear()
        bar_u = st.progress(0.0)
        for k_, f_ in enumerate(ups):
            bar_u.progress(k_ / len(ups), text=f_.name)
            try:
                out_, rep_ = hupd.update(f_.getvalue(), clean_name(f_.name))
                ures[f_.name] = dict(data=out_, rep=rep_)
            except Exception as e:                           # noqa
                ures[f_.name] = dict(error=f'{type(e).__name__}: {e}')
        bar_u.progress(1.0, text='done')
    if ures:
        zb = io.BytesIO()
        with zipfile.ZipFile(zb, 'w', zipfile.ZIP_DEFLATED) as z:
            for n_, r_ in ures.items():
                if 'data' in r_:
                    z.writestr(r_['rep']['filename'], r_['data'])
        st.download_button(t('dl_all'), zb.getvalue(), file_name=f'HOKA_internal_{dt.date.today():%m%d%Y}.zip',
                           type='primary', use_container_width=True, key='upd_zip')
        for n_, r_ in ures.items():
            if 'error' in r_:
                st.error(f'{n_}: {r_["error"]}')
                continue
            rp = r_['rep']
            icon = '🔴' if rp['notes'] or rp['cif_issues'] else '✅'
            with st.expander(f"{icon} {rp['filename']} — {rp['tab']}", expanded=len(ures) == 1):
                for b_ in rp['blocks']:
                    if b_['prev']:
                        st.write(f"**{b_['variant']}** · {t('upd_prev')}: `{b_['prev']}`")
                    else:
                        st.warning(f"{b_['variant']} · {t('upd_none')}")
                    if b_['gone']:
                        st.warning(t('upd_gone') + ': ' + ', '.join(b_['gone']))
                if rp.get('open0'):
                    st.error(t('upd_open0') + ': ' + ', '.join(rp['open0']))
                if rp['notes']:
                    st.markdown(f"**{t('upd_notes')}**")
                    st.dataframe(pd.DataFrame([{'': x['variant'], 'Cell': x['cell'], 'Item': x['item'],
                                                'REMARK (before)': x['before'], '🔴 +': x['note']} for x in rp['notes']]),
                                 hide_index=True, use_container_width=True)
                elif any(b_['prev'] for b_ in rp['blocks']):
                    st.success(t('upd_same'))
                if rp['skipped']:
                    st.caption(t('upd_skip') + ': ' + ', '.join(c for c, _ in rp['skipped']))
                if rp['cif_issues']:
                    st.error(t('cif_bad'))
                    st.dataframe(pd.DataFrame([{t('cif_type'): x['type'], 'Item': x['detail'], 'Price': x.get('price'),
                                                'CIF': x['cif'], t('cif_usual'): x['seen'],
                                                t('cif_why'): t('cif_' + x['why'])} for x in rp['cif_issues']]),
                                 hide_index=True, use_container_width=True)
                st.download_button('⬇️ ' + rp['filename'], r_['data'], file_name=rp['filename'], key=f'upd_dl_{n_}')

# ======================= FOB HISTORY =======================
with tab_hist:
    stor = fstore.Store(_secrets())
    if stor.kind == 'local':
        st.info(t('store_local'))
    try:
        hist = stor.load()
    except Exception as e:                                   # noqa
        st.error(f'{type(e).__name__}: {e}')
        hist = pd.DataFrame(columns=fstore.COLS)

    # ---- import old buyer CBS ----
    with st.expander(t('import'), expanded=hist.empty):
        st.caption(t('import_hint'))
        old = st.file_uploader(t('import'), type=['xlsx'], accept_multiple_files=True, key='hist_up',
                               label_visibility='collapsed')
        if old:
            recs, bad = [], []
            for f in old:
                try:
                    recs += fstore.records_from_buyer_file(f.name, f.getvalue())
                except Exception as e:                       # noqa
                    bad.append(f'{f.name}: {e}')
            for b in bad:
                st.warning(b)
            if recs:
                ed = st.data_editor(pd.DataFrame(recs), key='hist_ed', hide_index=True, use_container_width=True,
                                    num_rows='dynamic',
                                    column_config={'fob': st.column_config.NumberColumn(format='$%.2f'),
                                                   'color': st.column_config.SelectboxColumn(options=['SOLID', 'HEATHER'])})
                if st.button(t('save'), type='primary'):
                    a, rp = stor.upsert(ed.to_dict('records'))
                    st.success(t('saved').format(a=a, r=rp))
                    st.rerun()

    if hist.empty:
        st.info(t('empty'))
    else:
        c1, c2, c3 = st.columns([1, 1, 1])
        seasons = sorted(hist['season'].dropna().unique(), reverse=True)
        sel_s = c1.multiselect(t('f_season'), seasons, default=seasons[:1])
        stages = sorted(hist[hist['season'].isin(sel_s)]['stage'].dropna().unique()) if sel_s else []
        sel_g = c2.multiselect(t('f_stage'), stages, default=stages[-1:] if stages else [])
        q = c3.text_input(t('f_style'))
        view = hist[hist['season'].isin(sel_s) & hist['stage'].isin(sel_g)]
        if q:
            view = view[view['style'].str.contains(q.strip())]
        view = fstore.sort_table(view)
        cmp_ = st.toggle(t('compare'), value=len(sel_g) > 1 or len(sel_s) > 1)
        if cmp_ and (len(sel_g) > 1 or len(sel_s) > 1):
            vv = view.assign(step=view['season'] + ' ' + view['stage'])
            piv = vv.pivot_table(index=['style', 'color'], columns='step', values='fob', aggfunc='last')
            piv = piv[sorted(piv.columns, key=lambda x: (x[1:3], x[0] != 'S', x[4:]))]
            if len(piv.columns) >= 2:
                piv['Δ'] = piv.iloc[:, -1] - piv.iloc[:, 0]
            st.dataframe(piv.style.format('${:.2f}', na_rep=''), use_container_width=True)
        else:
            show = view[['season', 'stage', 'style', 'color', 'fob', 'date']].rename(columns={
                'season': 'SEASON', 'stage': 'STAGE', 'style': 'STYLE#', 'color': 'COLOR', 'fob': 'FOB (USD)', 'date': 'DATE'})
            st.dataframe(show, hide_index=True, use_container_width=True,
                         column_config={'FOB (USD)': st.column_config.NumberColumn(format='$%.2f')})
        if not view.empty:
            st.markdown(f"**{t('chart')}**")
            ch = view.assign(label=view['style'] + ' ' + view['color'].str[:1], step=view['season'] + ' ' + view['stage'])
            st.bar_chart(ch.pivot_table(index='label', columns='step', values='fob', aggfunc='last'), stack=False)
            st.download_button(t('dl_xlsx'), fstore.report_xlsx(view),
                               file_name=f"HOKA_FOB_{'_'.join(sel_s)}_{'_'.join(sel_g)}.xlsx")

