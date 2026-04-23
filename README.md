# TSL Password Batch Generator

## 功能
- 讀取 A.xlsx
- 逐筆將 KEY 傳入 TSL_password_generator.exe
- 擷取 Admin Password 與 Power User Password
- 產出 B.xlsx

## 預設欄位
- Device ID
- KEY

## 輸出欄位
- Device ID
- KEY
- Admin Password
- Power User Password
- Status
- Message

## 本機執行
```bash
python main.py --input A.xlsx --exe TSL_password_generator.exe --output B.xlsx
```

## GitHub Actions 產出 Windows EXE
把以下檔案放進 repository：
- main.py
- requirements.txt
- TSL_password_generator.exe
- A.xlsx（測試用，可選）
- .github/workflows/build-windows-exe.yml

Push 之後到 GitHub Actions 執行 workflow，就會產生 zip artifact。
