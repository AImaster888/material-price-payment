# material-price-payment — 專案規則

估驗計價逐項核算工具。核心引擎在 `scripts/verify_payment.py`，顯示文字渲染在
`scripts/render_excel.py`。改這兩支腳本前先讀完本檔，改完一定要跑
`python tests/run_tests.py` 全綠才算完成。

## 鐵則

1. **AI 不重算任何數字**。所有數字都由 `verify_payment.py` 用 openpyxl 讀出、用
   Python 算出來；Claude／其他 AI 只負責讀腳本輸出、翻譯成人話、決定下一步、
   在使用者要求時整理成報告。絕對不要在對話裡口算或憑印象覆核金額。

2. **比對要看「畫面顯示文字」，不能只看底層數值**（見 `docs/FORMATS.md` 的
   「為什麼有 display-text 檢查」）。兩份不同機關/不同人員製作的估驗計價檔，
   儲存格顯示格式（小數位數等）常常不一樣，底層數值相等不代表審查時畫面看
   起來一樣。`check_display_text_cross()` 就是為了這件事存在的，不要因為
   「反正數學上一樣」就把這個檢查拿掉或忽略它的結果。

3. **欄位版型用自動偵測，不要寫死欄位字母**。不同機關的範本，「原契約／前期
   累計／本期完成／累計至本期」這幾個群組標題的欄位位置常常不一樣（甚至同一
   個案子的第一期跟第二期都可能不同，例如「原契約」在某些檔案寫成「原派工」
   ——見 `GROUP_SYNONYMS`）。遇到偵測不到的新版型：先跑
   `SheetLayout(wf).describe()` 看看抓到什麼，缺什麼群組/子欄位就回報使用者，
   不要用固定欄位字母硬改腳本去遷就單一檔案。

4. **跨期比對的檔案順序是「前一期在前、本期在後」**，不可顛倒。
   `check_cross_period(wv_a, layout_a, wv_b, layout_b)` 的語意是
   「A的累計至本期 應等於 B的前期累計」，call 的時候 A 一定是較早那一期。

5. **容錯值（tolerance）不要隨便放大**。複價比對預設誤差 < 0.5（元），數量比對
   預設 < 0.0001，這是為了抓「捨入到小數點最後一位」等級的落差，不是抓不到
   問題就放大容忍度蓋過去。真的需要調整，要先問使用者為什麼、寫進
   `docs/FORMATS.md`。

## 改動流程

- 改 `scripts/verify_payment.py` 或 `scripts/render_excel.py` → 跑
  `python tests/run_tests.py`，全綠才算完成。
- 新增一種檢查（今天沒有的第 6 種模式）→ 先寫測試案例到 `tests/run_tests.py`
  （用假物件，不要依賴真的 xlsx 檔案跟 LibreOffice，因為 openpyxl 存檔後
  data_only 讀不到公式快取值，這個環境也跑不動 LibreOffice 重算）。
- 解析不了的新版型 → 照 `docs/FORMATS.md` 的「新增群組同義詞」流程處理，不要
  手工抄數字應急、不要為了單一案子的怪版型犧牲通用性。

## 隱私

- 這個 repo 是拿來分享工具本身的，**不要把任何真實案件的 xlsx 檔案、核算報告
  commit 進去**（`.gitignore` 已經擋掉 `*.xlsx`／`*核算報告*`／`*比對表*`，
  但新增規則時要留意別不小心繞過）。
- 案件筆記（含真實契約數字）另外存在 gitignore 掉的檔案裡，不要複製進
  README.md 或任何會被 commit 的文件。
