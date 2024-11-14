# lmsV1Tools

- 月讀 77 折工具

```bash
go run coupon.go
```
依序輸入資料後可產生以下四個檔案
noRow.csv => 代表 excel 中的 row 查不到對應資料
overOneRowFile.csv => 代表有兩張以上的 77 折，需要檢查該如何處理
tradeSearch.sql => 查詢交易紀錄的 sql
update.sql => 更新用的 sql