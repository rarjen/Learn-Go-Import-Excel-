package main

import (
	"database/sql"
	"excel-import-export/utils"
	"log"

	"github.com/gin-gonic/gin"
	_ "github.com/go-sql-driver/mysql"
)

var db *sql.DB

func main() {
	gin.SetMode(gin.ReleaseMode)
	var err error
	db, err = sql.Open("mysql", "root@tcp(localhost:3306)/db_import_go")
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()

	r := gin.Default()

	r.POST("/api/import", utils.ImportHandler(db))
	r.GET("/api/export", utils.ExportHandler(db))

	port := ":9080"

	// Log the message to the terminal
	log.Printf("Server is running on port %s\n", port)

	// Start the server
	err = r.Run(port)
	if err != nil {
		log.Fatal(err)
	}
}
