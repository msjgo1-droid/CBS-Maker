# HOKA CBS Builder (Streamlit)

Internal CBS → (클릭 한 번) 정돈된 Internal + 바이어 제출용 HOKA CBS. 여러 스타일 동시 처리(스타일번호 7자리로 자동 묶음). UI: 한국어 / English / Tiếng Việt

## 실행
```
pip install -r requirements.txt
streamlit run app.py
```
배포: GitHub repo(private 권장)에 올리고 Streamlit Community Cloud에서 `app.py` 지정.
`packages.txt`(libreoffice-draw)는 Internal 스케치가 WMF/EMF일 때 PNG 변환용 — 없으면 해당 스타일만 경고 후 스케치 생략.

## 업로드 조합
| 경우 | Internal | 바이어 template | 이전 제출본 | MNMK |
|---|---|---|---|---|
| CO 스타일 | ✅ | 새 template (없으면 이전 제출본을 베이스로) | ✅ | – |
| 신규 / MNMK 반영 | ✅ | ✅ | – | ✅ |
| CO + MNMK가 틀린 경우 | ✅ | ✅ | ✅ | ✅ |

## 규칙 (요약)
- Internal = 항상 정답. 맨 앞 탭 = 최신. `CBS m/d $금액` 메모가 FOB와 같으면 제출 블록.
- 바이어 template: Net Usage = internal 바이어쪽 yy, %Loss = LOSS−1, Price ≠ internal 단가면 Price Override,
  Freight Override = CIF, Labor = CM, OH/Profit/Transportation = SUM(Material:Labor) × 5%/5%/3%.
- 항목 매칭: 자재코드 → 이름 → 단어 유사도 → 바이어 Price/Loss 숫자 비교. 화면에서 수동 수정 가능.
- 같은 원단이 바이어 template 여러 행(Shell 1 / Contrast 1 등)에 있고 Internal은 1줄일 때: Internal YY/Detail/Remark(또는 바로 아래 줄)에 `a+b`가 있으면 그 비율로 나눔 → 없으면 이전 제출 CBS가 나눠져 있었을 때 그 비율 → 둘 다 없으면 첫 행에 전부, 나머지 0. 원단(마커) 시트는 합계 yy 기준.
- template에 없는 항목은 해당 섹션 끝에 행 추가 (Interlining 제외).
- Solid/Heather: 이전 제출본에 둘 다 있으면 둘 다 생성, 반대쪽 variant 행은 0.
- 원단(마커) 시트: MNMK가 최신(날짜·폭·효율·Panel/Garment Q'ty·사이즈·사진), 마커 길이(Y/IN)는
  internal yy로 역산 → `((IN/36)+Y)/Garment Q'ty = yy`. 한 원단이 여러 시트면 보조 시트(BIAS CUT 등)는 그대로, 메인 시트 = yy − 보조 yy.
- 사이즈 블록: 남 M / 여 S. MNMK 숨김 시트·Internal 숨김 시트는 경고.
- 파일명: 이전 바이어 파일명 그대로, 공장코드 HJ → PIVNVN, 날짜 MMDDYYYY(작성일).
- Internal 서식 정리: 최종가격(internal 합계·open FOB·상단 FOB)만 소수점 2자리, 나머지 숫자는 최대 4자리. XML 직접 수정이라 이미지(WMF 포함) 그대로 유지.
- 바이어 CBS: 상단 FOB 왼쪽에 Internal 가먼트 스케치 삽입, FOB는 소수점 2자리, OH/Profit/Transportation 비율은 5%/5%/3%(오른쪽 정렬).
- Self(body) 마커는 1pc / 1way(lengthwise) 원칙 — 마커 사진에서 같은 색(=가먼트 1벌) 패널 묶음이 길이 방향으로 모였는지 판정, lengthwise 면적 75% 미만이면 실행 시 맨 위 빨간 경고 + 표시된 사진.
- 정돈된 Internal: REMARK 칸에 단계(섹션)별 가격 합계(=SUM)가 없으면 추가 — 원단 / Special Trim / Embellishment / Thread / Labels / Packing / Profit·OH·Finance. 이미 있는 합계(여러 섹션을 묶은 것 포함)는 절대 건드리지 않음(Profit 수식이 참조하기 때문).
- 정돈된 Internal: 사용된 블록의 open(바이어) FOB 오른쪽 칸에 "CBS 실행일 $바이어FOB" 기록 (예: CBS 11/9 $23.76).

## 알려진 제한
- 바이어 template/이전 제출본에 WMF/EMF 이미지가 있으면 결과 파일에서 빠짐(경고 표시).
- 수동 매칭은 세션 동안만 기억.

## 📊 FOB 기록 (시즌 · 단계별 가격 조회)
- CBS 만들기 결과에서 **💾 이 결과를 FOB 기록에 저장** → 시즌 / 단계 / 스타일 / 컬러(SOLID·HEATHER) / FOB / 날짜 저장.
  같은 시즌+단계+스타일+컬러는 최신 값으로 덮어씀.
- **📊 FOB 기록** 탭: 예전에 바이어에게 나간 CBS 여러 개를 올려 한 번에 등록(파일명·시트에서 자동 추출, 표에서 수정 가능,
  공란 template은 자동 제외) → 시즌·단계 선택 → 표 / 차트 / 엑셀 다운로드, 여러 단계·시즌 선택 시 열로 나란히 비교(Δ).

### 저장소 설정 (Streamlit Cloud에서는 필수)
Streamlit Cloud는 앱이 재시작되면 서버 파일이 지워져서, 기록은 **별도 private GitHub repo의 CSV**에 저장해요.
1. GitHub에 빈 private repo 생성 (예: `hoka-cbs-data`) — 앱 repo와 다른 repo여야 저장할 때마다 앱이 재배포되지 않음.
2. GitHub → Settings → Developer settings → Fine-grained token: 그 repo만, **Contents: Read and write**.
3. Streamlit Cloud → 앱 → Settings → Secrets 에 입력:
```
[fob_store]
github_token = "github_pat_..."
github_repo  = "내아이디/hoka-cbs-data"
path         = "fob_history.csv"
branch       = "main"
```
설정이 없으면 서버 임시 폴더(data/fob_history.csv)에 저장되고 화면에 경고가 떠요 (내 PC에서 실행할 때는 그대로 사용 가능).
