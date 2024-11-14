package main

import (
	"database/sql"
	"errors"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	_ "github.com/denisenkom/go-mssqldb"
	"github.com/xuri/excelize/v2"

	"lmsV1Tools/database"
)

type ExcelRowData struct {
	id                 string // 編號
	storeCode          string // 門市代碼
	memberName         string // 會員姓名
	memberCode         string // 會員編號
	tradeDate          string // 交易日期
	invoiceNumber      string // 發票號碼
	tradeMachineNumber string // 交易機號
	tradeNumber        string // 交易序號
}

func main() {
	//overOneMonth()
	thisMonth()
}

func overOneMonth() {
	ignoreCount := 0
	// 初始化資料庫連接

	db := database.GetDB()

	var year int         // 處理年份
	var month int        // 處理月份
	var staffCode string // 誠品員工編號
	fmt.Println("請輸入處理年份：")
	_, err := fmt.Scanln(&year)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("請輸入處理月份：")
	_, err = fmt.Scanln(&month)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("請輸入您的員工編號：")
	_, err = fmt.Scanln(&staffCode)
	if err != nil {
		fmt.Println(err)
		return
	}
	if len(os.Args) < 2 {
		fmt.Println("No arguments provided")
		return
	}
	filename := os.Args[1]
	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Get all the rows in the first sheet.
	sheets := f.GetSheetList()
	rows, err := f.GetRows(sheets[0])
	if err != nil {
		fmt.Println(err)
		return
	}

	tradeSearchSqlFile, err := os.Create("tradeSearch.sql")
	if err != nil {
		fmt.Println("Error creating file:", err)
	}
	defer tradeSearchSqlFile.Close()

	updateSqlFile, err := os.Create("update.sql")
	if err != nil {
		fmt.Println("Error creating file:", err)
	}
	defer updateSqlFile.Close()

	noRowFile, err := os.Create("noRow.csv")
	if err != nil {
		fmt.Println("Error creating file:", err)
	}
	defer noRowFile.Close()

	// 獲得處理月份的第一天
	firstDay := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.UTC)
	// 獲得處理月份的最後一天
	lastDay := firstDay.AddDate(0, 1, -1)
	// 獲取當前時間
	now := time.Now()

	// 格式化成 yyyy-mm-dd 的形式
	today := now.Format("2006-01-02")

	firstDayFormatted := firstDay.Format("2006/01/02")
	lastDayFormatted := lastDay.Format("2006/01/02")

	tradeSearchSqlTemplate := "SELECT SL_DATE, S_NO, ID_NO, SL_NO FROM CUSL_H WHERE C_NO='%s' and SL_NO='%s'"

	updateSqlTemplate := "UPDATE COUPON_20170214 SET re_date='%s',re_s_no='%s',re_id_no='%s',re_sl_no='%s',CP_STATUS='Y',UPD_DATE='%s',UPD_USER='%s' WHERE CP_NO='%s' AND CP_STATUS='A' AND C_NO = '%s';"
	//"UPDATE COUPON_20170214 SET re_date='%s',re_s_no='B089',re_id_no='001',re_sl_no='08210001',CP_STATUS='Y',UPD_DATE='%s',UPD_USER='808648' WHERE CP_NO='8882154690211200891004' and CP_STATUS='A'"

	processCount := 0
	for _, row := range rows {
		// 如果 row[0] 不是一個數字，跳過處理
		if _, err := strconv.Atoi(row[0]); err != nil {
			continue
		}
		processCount++
		excelRowData := ExcelRowData{
			id:                 strings.TrimSpace(row[0]),
			storeCode:          strings.TrimSpace(row[1]),
			memberName:         strings.TrimSpace(row[2]),
			memberCode:         strings.TrimSpace(row[3]),
			tradeDate:          strings.TrimSpace(row[4]),
			invoiceNumber:      strings.TrimSpace(row[5]),
			tradeMachineNumber: strings.TrimSpace(row[6]),
			tradeNumber:        strings.TrimSpace(row[7]),
		}
		fmt.Println(excelRowData)

		tradeSearchQuery := fmt.Sprintf(tradeSearchSqlTemplate, excelRowData.memberCode, excelRowData.tradeNumber)
		row := db.QueryRow(tradeSearchQuery)
		var SL_DATE string
		var S_NO string
		var ID_NO string
		var SL_NO string
		err := row.Scan(&SL_DATE, &S_NO, &ID_NO, &SL_NO)
		tradeDateSearchSqlTemplate := "SELECT SL_DATE, S_NO, ID_NO, SL_NO FROM CUSL_H WHERE C_NO='%s' and SL_DATE='%s'"
		if errors.Is(err, sql.ErrNoRows) {
			// 如果透過交易序號 + 員編找不到，改用日期 + 員編找找看
			dateStr := excelRowData.tradeDate
			// 定義輸入的日期格式
			layout := "1月2日"
			parsedDate, err := time.Parse(layout, dateStr)
			if err != nil {
				fmt.Println("解析日期錯誤:", err)
				return
			}

			// 使用指定的年份 (例如 2024)
			year := 2024
			finalDate := time.Date(year, parsedDate.Month(), parsedDate.Day(), 0, 0, 0, 0, time.Local)

			// 格式化為 "2024/9/29" 格式
			formattedDate := finalDate.Format("2006/01/02")
			tradeDateSearchQuery := fmt.Sprintf(tradeDateSearchSqlTemplate, excelRowData.memberCode, formattedDate)
			row = db.QueryRow(tradeDateSearchQuery)
			err = row.Scan(&SL_DATE, &S_NO, &ID_NO, &SL_NO)
			if errors.Is(err, sql.ErrNoRows) {
				noRowFile.WriteString(tradeDateSearchQuery + "\n")
				continue
			}
		} else if err != nil {
			log.Fatal(err)
		}

		var CP_NO string
		cpNoQueryTemplate := "select CP_NO from COUPON_20170214 c inner join GIFT g on c.GIFT_NO=g.GIFT_NO where C_NO='%s' and CP_STATUS='A' and CP_START_DATE>='%s' and CP_START_DATE<='%s' and c.GIFT_NO = '20170214'"
		cpNoQuery := fmt.Sprintf(cpNoQueryTemplate, excelRowData.memberCode, firstDayFormatted, lastDayFormatted)
		fmt.Println(cpNoQuery)
		time.Sleep(10000)
		cpRow := db.QueryRow(cpNoQuery)
		cpErr := cpRow.Scan(&CP_NO)
		if errors.Is(cpErr, sql.ErrNoRows) {
			// 代表他的 COUPON_20170214.CP_STATUS 是 Y，已經用過，不需要 update
			ignoreCount++
			continue
		} else if cpErr != nil {
			log.Fatal(cpErr)
		}

		updateSqlQuery := fmt.Sprintf(updateSqlTemplate, SL_DATE, S_NO, ID_NO, SL_NO, today, staffCode, CP_NO, excelRowData.memberCode)
		_, err = updateSqlFile.WriteString(updateSqlQuery + "\n")
		if err != nil {
			fmt.Println("Error writing file:", err)
		}
	}
	fmt.Println("總處理的數量：", processCount)
	fmt.Println("不需要處理的數量：", ignoreCount)
}

func thisMonth() {
	ignoreCount := 0
	// 初始化資料庫連接

	db := database.GetDB()

	var year int         // 處理年份
	var month int        // 處理月份
	var staffCode string // 誠品員工編號
	fmt.Println("請輸入處理年份：")
	_, err := fmt.Scanln(&year)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("請輸入處理月份：")
	_, err = fmt.Scanln(&month)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("請輸入您的員工編號：")
	_, err = fmt.Scanln(&staffCode)
	if err != nil {
		fmt.Println(err)
		return
	}
	if len(os.Args) < 2 {
		fmt.Println("No arguments provided")
		return
	}
	filename := os.Args[1]
	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Get all the rows in the first sheet.
	sheets := f.GetSheetList()
	rows, err := f.GetRows(sheets[0])
	if err != nil {
		fmt.Println(err)
		return
	}

	tradeSearchSqlFile, err := os.Create("tradeSearch.sql")
	if err != nil {
		fmt.Println("Error creating file:", err)
	}
	defer tradeSearchSqlFile.Close()

	updateSqlFile, err := os.Create("update.sql")
	if err != nil {
		fmt.Println("Error creating file:", err)
	}
	defer updateSqlFile.Close()

	noRowFile, err := os.Create("noRow.csv")
	if err != nil {
		fmt.Println("Error creating file:", err)
	}
	defer noRowFile.Close()

	overOneRowFile, err := os.Create("overOneRowFile.csv")
	if err != nil {
		fmt.Println("Error creating file:", err)
	}
	defer overOneRowFile.Close()

	// 獲得處理月份的第一天
	firstDay := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.UTC)
	// 獲得處理月份的最後一天
	lastDay := firstDay.AddDate(0, 1, -1)
	// 獲取當前時間
	now := time.Now()

	// 格式化成 yyyy-mm-dd 的形式
	today := now.Format("2006-01-02")

	firstDayFormatted := firstDay.Format("2006/01/02")
	lastDayFormatted := lastDay.Format("2006/01/02")

	searchCountSqlTemplate := "select count(*) as ROW_COUNT from COUPON WHERE C_NO = '%s' AND GIFT_NO = '20170214' AND CP_STATUS = 'A' AND CP_START_DATE>='%s' AND CP_START_DATE<='%s'"
	searchTemplate := "select 1 from COUPON WHERE C_NO = '%s' AND GIFT_NO = '20170214' AND CP_STATUS = 'A' AND CP_START_DATE>='%s' AND CP_START_DATE<='%s'"
	updateSqlTemplate := "UPDATE COUPON SET CP_STATUS='Y',UPD_DATE='%s',UPD_USER='%s' WHERE CP_STATUS='A' AND C_NO = '%s' AND CP_START_DATE>='%s' AND CP_START_DATE<='%s' AND GIFT_NO = '20170214';"

	processCount := 0
	for _, row := range rows {
		// 如果 row[0] 不是一個數字，跳過處理
		if _, err := strconv.Atoi(row[0]); err != nil {
			continue
		}
		processCount++
		excelRowData := ExcelRowData{
			id:                 strings.TrimSpace(row[0]),
			storeCode:          strings.TrimSpace(row[1]),
			memberName:         strings.TrimSpace(row[2]),
			memberCode:         strings.TrimSpace(row[3]),
			tradeDate:          strings.TrimSpace(row[4]),
			invoiceNumber:      strings.TrimSpace(row[5]),
			tradeMachineNumber: strings.TrimSpace(row[6]),
			tradeNumber:        strings.TrimSpace(row[7]),
		}
		fmt.Println(excelRowData)

		var custExist string
		err := db.QueryRow(fmt.Sprintf(
			"SELECT 1 FROM CUST WHERE C_NO = '%s'",
			excelRowData.memberCode,
		)).Scan(&custExist)
		if err != nil {
			noRowFile.WriteString(fmt.Sprintf(
				"SELECT 1 FROM CUST WHERE C_NO = '%s'",
				excelRowData.memberCode,
			) + "\n")
			ignoreCount++
			continue
		}

		var rowCount string
		couponRowCountSearchQuery := fmt.Sprintf(searchCountSqlTemplate, excelRowData.memberCode, firstDayFormatted, lastDayFormatted)
		err = db.QueryRow(couponRowCountSearchQuery).Scan(&rowCount)
		if err != nil {
			overOneRowFile.WriteString(couponRowCountSearchQuery + "\n")
			ignoreCount++
			continue
		}

		couponSearchQuery := fmt.Sprintf(searchTemplate, excelRowData.memberCode, firstDayFormatted, lastDayFormatted)
		var dummy int
		err = db.QueryRow(couponSearchQuery).Scan(&dummy)
		if errors.Is(err, sql.ErrNoRows) {
			noRowFile.WriteString(couponSearchQuery + "\n")
			ignoreCount++
			continue
		} else if err != nil {
			log.Fatal(err)
		}

		updateSqlQuery := fmt.Sprintf(updateSqlTemplate, today, staffCode, excelRowData.memberCode, firstDayFormatted, lastDayFormatted)
		_, err = updateSqlFile.WriteString(updateSqlQuery + "\n")
		if err != nil {
			fmt.Println("Error writing file:", err)
		}
	}
	fmt.Println("總處理的數量：", processCount)
	fmt.Println("不需要處理的數量：", ignoreCount)
}
