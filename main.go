package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/AlecAivazis/survey/v2"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"

	"lmsV1Tools/database"
	userService "lmsV1Tools/services/user_service"
)

var rootCmd = &cobra.Command{
	Use:   "cli-tool",
	Short: "LMS CLI 工具",
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func init() {
	rootCmd.AddCommand(getUserEmailByExcel)
}

type CUST struct {
	MemberCode string `gorm:"column:C_NO"`
	Email      string `gorm:"column:C_EMAIL"`
}

var getUserEmailByExcel = &cobra.Command{
	Use:   "get-user-email-excel",
	Short: "使用指定 excel 檔案（包含 C_NO）來產生含有 email 的檔案",
	Run: func(cmd *cobra.Command, args []string) {
		var dir string
		var users []CUST
		prompt := &survey.Input{
			Message: "請輸入資料夾路徑:",
		}
		survey.AskOne(prompt, &dir)

		files, err := scanDirectory(dir)
		if err != nil {
			log.Fatalf("掃描資料夾失敗: %v", err)
		}

		var selectedFile string
		promptFile := &survey.Select{
			Message: "選擇檔案:",
			Options: files,
		}
		survey.AskOne(promptFile, &selectedFile)

		fmt.Printf("選擇了檔案: %s\n", selectedFile)

		db := database.GetDB()
		userCustomerNumbers := userService.GetUserCustomerNumbersByExcelFile(selectedFile)

		// 建立寫入資料用的 excel 檔案
		outputFile := excelize.NewFile()
		defer func() {
			if err := outputFile.Close(); err != nil {
				fmt.Println(err)
			}
		}()

		outputFile.NewSheet("Sheet1")

		batchSize := 1000 // 設定批次大小
		excelRowCount := 2
		outputFile.SetSheetRow("Sheet1", "A1", &[]any{"會員編號", "Email"})
		for i := 0; i < len(userCustomerNumbers); i += batchSize {
			fmt.Printf("當前處理 %d 到 %d\n", i, i+batchSize)
			end := i + batchSize
			if end > len(userCustomerNumbers) {
				end = len(userCustomerNumbers)
			}

			batch := userCustomerNumbers[i:end]
			fmt.Printf("%T\n", batch)
			fmt.Println(batch)
			result := db.QueryRow("SELECT C_NO, C_EMAIL FROM CUST WHERE C_NO IN ?", batch).Scan(&users)

			if result.Error != nil {
				fmt.Println("Error querying users:", result.Error)
				return
			}

			for _, user := range users {
				writeRow := []any{user.MemberCode, user.Email}
				cell := fmt.Sprintf("A%d", excelRowCount)
				outputFile.SetSheetRow("Sheet1", cell, &writeRow)
				excelRowCount++
			}
		}

		if err = outputFile.SaveAs("files/userEmails.xlsx"); err != nil {
			fmt.Println(err)
		}
	},
}

func scanDirectory(dir string) ([]string, error) {
	var files []string
	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if !info.IsDir() {
			files = append(files, path)
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return files, nil
}
