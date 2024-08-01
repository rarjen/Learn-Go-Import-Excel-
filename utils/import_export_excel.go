// utils/import_export.go
package utils

import (
	"database/sql"
	"fmt"
	"net/http"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gin-gonic/gin"
)

func ImportHandler(db *sql.DB) gin.HandlerFunc {
	return func(c *gin.Context) {
		file, err := c.FormFile("file")
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
			return
		}

		// Read data from Excel files
		fileContent, err := file.Open()
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
			return
		}
		defer fileContent.Close()

		xlsx, err := excelize.OpenReader(fileContent)
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
			return
		}

		// Start a transaction
		tx, err := db.Begin()
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
			return
		}

		// Extract data from an Excel file and insert it into a database
		rows := xlsx.GetRows("Sheet1")
		for i, row := range rows {
			if i == 0 {
				// Skip header row
				continue
			}

			// Check for empty fields
			for _, cell := range row {
				if cell == "" {
					tx.Rollback() // Rollback the transaction
					c.JSON(http.StatusBadRequest, gin.H{"error": fmt.Sprintf("There is an empty field on row %d", i+1)})
					return
				}
			}

			code := row[0]
			name := row[1]
			latitude := row[2]
			longitude := row[3]
			address := row[4]
			city := row[5]
			operation_hour := row[6]

			// Check if the code already exists in the database
			var existingID int
			err := tx.QueryRow("SELECT id FROM store_locations WHERE code = ?", code).Scan(&existingID)
			if err != nil && err != sql.ErrNoRows {
				tx.Rollback() // Rollback the transaction
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			if existingID != 0 {
				// Code already exists, return error message
				tx.Rollback() // Rollback the transaction
				c.JSON(http.StatusConflict, gin.H{"error": fmt.Sprintf("Data with code %s already exists", code)})
				return
			}

			// Save data to the database
			_, err = tx.Exec("INSERT INTO store_locations (code, name, latitude, longitude, address, city, operation_hour) VALUES (?, ?, ?, ?, ?, ?, ?)", code, name, latitude, longitude, address, city, operation_hour)
			if err != nil {
				tx.Rollback() // Rollback the transaction
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}
		}

		// Commit the transaction
		if err := tx.Commit(); err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
			return
		}

		c.JSON(http.StatusOK, gin.H{"message": "Data imported successfully"})
	}
}

func ExportHandler(db *sql.DB) gin.HandlerFunc {
	return func(c *gin.Context) {
		// Retrieve data from the database
		rows, err := db.Query("SELECT id, code, name, latitude, longitude, address, city, operation_hour FROM store_locations")
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
			return
		}
		defer rows.Close()

		// Create an Excel file
		xlsx := excelize.NewFile()
		sheetName := "Sheet1"
		xlsx.SetSheetName("Sheet1", sheetName)

		// Add headers
		xlsx.SetCellValue(sheetName, "A1", "Code")
		xlsx.SetCellValue(sheetName, "B1", "Name")
		xlsx.SetCellValue(sheetName, "C1", "Latitude")
		xlsx.SetCellValue(sheetName, "D1", "Longitude")
		xlsx.SetCellValue(sheetName, "E1", "Address")
		xlsx.SetCellValue(sheetName, "F1", "City")
		xlsx.SetCellValue(sheetName, "G1", "Operation Hour")

		// Fill data from the database into an Excel file
		rowIndex := 2
		for rows.Next() {
			var id int
			var code, name, latitude, longitude, address, city, operation_hour string

			err := rows.Scan(&id, &code, &name, &latitude, &longitude, &address, &city, &operation_hour)
			if err != nil {
				c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
				return
			}

			xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), code)
			xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), name)
			xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowIndex), latitude)
			xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowIndex), longitude)
			xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowIndex), address)
			xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowIndex), city)
			xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowIndex), operation_hour)

			rowIndex++
		}

		// Set header response
		c.Header("Content-Transfer-Encoding", "binary")
		c.Header("Content-Disposition", "attachment; filename=store_locations_export.xlsx")
		c.Header("Content-Type", "application/octet-stream")

		// Write the Excel file to the response
		err = xlsx.Write(c.Writer)
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
			return
		}
	}
}
