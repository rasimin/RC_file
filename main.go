package main

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

const (
	basedir    = "."
	fileName   = "data2.csv"
	fileNamexl = "data.xlsx"
)

func main() {
	if len(os.Args) < 2 {
		log.Fatal("please provide a mode (csv or xls)")
	}

	mode := os.Args[1]

	if mode == "csv" {
		processCSV()
	} else if mode == "xls" {
		processXLS()
	} else {
		log.Fatal("invalid mode. Please use 'csv' or 'xls'")
	}
}

func processCSV() {
	csvFile, err := os.Open(basedir + "/" + fileName)
	if err != nil {
		log.Fatalf("failed creating file: %s", err)
	}
	defer csvFile.Close()

	reportFile, _ := os.Create(basedir + "/report.log")
	reportFileWriter := bufio.NewWriter(reportFile)
	_ = reportFileWriter.Flush()

	lines, err := csv.NewReader(csvFile).ReadAll()
	if err != nil {
		log.Fatalf("failed reading csv file: %s", err)
	}

	for i := 1; i < len(lines); i++ {
		line := lines[i]
		name, address := line[0], line[1]
		//if i == 2 {
		report := fmt.Sprintf("Konten baris ke %d: name: %s address: %s \n", i, name, address)
		//fmt.Println(report)
		_, _ = fmt.Fprintf(reportFileWriter, report)
		_ = reportFileWriter.Flush()
		//}
	}

}

func processXLS() {
	excelFile, err := excelize.OpenFile(basedir + "/" + fileNamexl)
	if err != nil {
		log.Fatalf("failed to open excel file: %s", err)
	}

	rows, err := excelFile.GetRows("Sheet1")
	if err != nil {
		log.Fatalf("failed to read rows: %s", err)
	}

	var names []string
	for _, row := range rows {
		var name string
		for i, colCell := range row {
			fmt.Print(colCell + ",")
			if i == 1 || i == 2 {
				name += colCell
			}
		}
		names = append(names, name)
		fmt.Println()
	}

	namesFile := excelize.NewFile()
	namesFileSheet1Index, _ := namesFile.NewSheet("Sheet1")
	namesFile.SetActiveSheet(namesFileSheet1Index)
	_ = namesFile.SetCellValue("Sheet1", "A1", "Name")
	for i, name := range names {
		_ = namesFile.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+2), name)
	}

	err = namesFile.SaveAs(basedir + "/names.xlsx")
	if err != nil {
		log.Fatalf("failed to save excel file: %s", err)
	}
}
