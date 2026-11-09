# HOKA CBS Builder (Streamlit)

Internal CBS → 바이어 제출용 HOKA CBS 자동 생성 + Internal 서식 정리. UI: KO / EN / VI

## 실행
```
pip install -r requirements.txt
streamlit run app.py
```
배포: GitHub repo(private 권장)에 올리고 Streamlit Community Cloud에서 `app.py` 지정.

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
- template에 없는 항목은 해당 섹션 끝에 행 추가 (Interlining 제외).
- Solid/Heather: 이전 제출본에 둘 다 있으면 둘 다 생성, 반대쪽 variant 행은 0.
- 원단(마커) 시트: MNMK가 최신(날짜·폭·효율·Panel/Garment Q'ty·사이즈·사진), 마커 길이(Y/IN)는
  internal yy로 역산 → `((IN/36)+Y)/Garment Q'ty = yy`. 한 원단이 여러 시트면 보조 시트(BIAS CUT 등)는 그대로, 메인 시트 = yy − 보조 yy.
- 사이즈 블록: 남 M / 여 S. MNMK 숨김 시트·Internal 숨김 시트는 경고.
- 파일명: 이전 바이어 파일명 그대로, 공장코드 HJ → PIVNVN, 날짜 MMDDYYYY(작성일).
- Internal 서식 정리: 단가·YY 최대 4자리, FOB·마진·Net·OH/Profit/Finance 2자리. XML 직접 수정이라 이미지(WMF 포함) 그대로 유지.

## 알려진 제한
- 바이어 template/이전 제출본에 WMF/EMF 이미지가 있으면 결과 파일에서 빠짐(경고 표시).
- 수동 매칭은 세션 동안만 기억.
