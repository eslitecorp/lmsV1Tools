package userService

import (
	"database/sql"
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type UserExcelRow struct {
	memberCode string
	email      sql.NullString
}

func GetUserCustomerNumbersByExcelFile(filepath string) []string {
	var customerNumbers []string
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		fmt.Errorf("error open file : %v", err)
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
		fmt.Errorf("error get rows: %v", err)
	}
	for _, row := range rows {
		if len(row) <= 3 {
			continue // 跳過長度不足的 row
		}

		// 嘗試將 row[3] 轉換為數字
		if _, err := strconv.Atoi(row[3]); err != nil {
			continue // 跳過無法轉換為數字的 row[3]
		}

		customerNumbers = append(customerNumbers, row[3])
	}

	return customerNumbers
}
