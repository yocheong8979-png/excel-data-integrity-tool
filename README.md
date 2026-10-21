# Excel 資料完整性分類工具

這是一個以 Streamlit 建立的互動式網頁應用程式。使用者可上傳 `.xlsx` 或 `.xls` Excel 檔，系統會檢查資料完整性問題，並產生包含「完整資料」與「不完整資料」兩個工作表的 Excel 檔。

## 功能

- 上傳 Excel 檔案
- 選擇要檢查的工作表
- 檢查空值、NULL、N/A、NA、None、unknown
- 檢查日期、電話、電郵格式
- 檢查年齡、薪金、消費金額、折扣率異常值
- 檢查分類欄位大小寫混亂
- 顯示完整資料與不完整資料統計
- 下載整理後的 Excel 檔

## 本機執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署到 Streamlit Community Cloud

1. 將 `app.py`、`requirements.txt` 與 `README.md` 上傳到 GitHub repository。
2. 前往 [Streamlit Community Cloud](https://share.streamlit.io/)。
3. 選擇「New app」。
4. 連結 GitHub 帳號並選擇 repository。
5. 設定：
   - Main file path：`app.py`
   - Python version：建議 3.11 或 3.12
6. 點選「Deploy」。

部署完成後，Streamlit 會提供一個公開網址，可直接讓使用者上傳 Excel 並下載整理結果。