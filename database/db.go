package database

import (
	"database/sql"
	"fmt"
	"log"
	"os"

	"github.com/joho/godotenv"
)

func GetDB() *sql.DB {
	err := godotenv.Load()
	if err != nil {
		log.Fatal("Error loading .env file")
	}
	// 設定資料庫連線資訊
	server := os.Getenv("DB_HOST")       // 資料庫伺服器位址
	port := os.Getenv("DB_PORT")         // 資料庫埠
	user := os.Getenv("DB_USER")         // 資料庫使用者
	password := os.Getenv("DB_PASSWORD") // 資料庫密碼
	database := os.Getenv("DB_DATABASE") // 資料庫名稱

	// 建立連線字串
	connString := fmt.Sprintf("server=%s;port=%s;user id=%s;password=%s;database=%s;",
		server, port, user, password, database)

	// 開始連線
	db, err := sql.Open("sqlserver", connString)
	if err != nil {
		log.Fatal("Open connection failed:", err.Error())
	}

	// 驗證連線是否成功
	err = db.Ping()
	if err != nil {
		log.Fatal("Ping failed:", err.Error())
	}

	return db
}
