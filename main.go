package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

//lint:file-ignore SA4006 trying out file-based linter directives

const (
	SheetName = "Expense Report"
)

var (
	expenseData = [][]interface{}{
		{"2022-04-10", "Flight", "Trip to San Fransisco", "", "", "$3,462.00"},
		{"2022-04-10", "Hotel", "Trip to San Fransisco", "", "", "$1,280.00"},
		{"2022-04-12", "Swags", "App launch", "", "", "$862.00"},
		{"2022-03-15", "Marketing", "App launch", "", "", "$7,520.00"},
		{"2022-04-11", "Event hall", "App launch", "", "", "$2,080.00"},
	}
)

func main() {
	var err error
	f := excelize.NewFile()
	index := f.NewSheet("Sheet1")
	f.SetActiveSheet(index)
	f.SetSheetName("Sheet1", SheetName)

	err = f.SetColWidth(SheetName, "A", "A", 6)
	err = f.SetColWidth(SheetName, "H", "H", 6)
	err = f.SetColWidth(SheetName, "B", "B", 12)
	err = f.SetColWidth(SheetName, "C", "C", 16)
	err = f.SetColWidth(SheetName, "D", "D", 13)
	err = f.SetColWidth(SheetName, "E", "E", 15)
	err = f.SetColWidth(SheetName, "F", "F", 22)
	err = f.SetColWidth(SheetName, "G", "G", 13)

	err = f.SetRowHeight(SheetName, 1, 12)

	err = f.MergeCell(SheetName, "A1", "H1")

	err = f.SetRowHeight(SheetName, 2, 25)
	err = f.MergeCell(SheetName, "B2", "D2")

	style, err := f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 20, Color: "6d64e8"}})
	err = f.SetCellStyle(SheetName, "B2", "D2", style)
	err = f.SetSheetRow(SheetName, "B2", &[]interface{}{"Gigashots Inc."})

	err = f.MergeCell(SheetName, "B3", "D3")
	err = f.SetSheetRow(SheetName, "B3", &[]interface{}{"3154 N Richardt Ave"})

	err = f.MergeCell(SheetName, "B4", "D4")
	err = f.SetSheetRow(SheetName, "B4", &[]interface{}{"Indianapolis, IN 46276"})

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	err = f.MergeCell(SheetName, "B5", "D5")
	err = f.SetCellStyle(SheetName, "B5", "D5", style)
	err = f.SetSheetRow(SheetName, "B5", &[]interface{}{"(317) 854-0398"})

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 32, Color: "2B4492", Bold: true}})
	err = f.MergeCell(SheetName, "B7", "G7")
	err = f.SetCellStyle(SheetName, "B7", "G7", style)
	err = f.SetSheetRow(SheetName, "B7", &[]interface{}{"Expense Report"})

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Color: "E25184", Bold: true}})
	err = f.MergeCell(SheetName, "B8", "C8")
	err = f.SetCellStyle(SheetName, "B8", "C8", style)
	err = f.SetSheetRow(SheetName, "B8", &[]interface{}{"09/04/00 - 09/05/00"})

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Bold: true}})
	err = f.SetCellStyle(SheetName, "B10", "G10", style)
	err = f.SetSheetRow(SheetName, "B10", &[]interface{}{"Name", "", "Employee ID", "", "Department"})
	err = f.MergeCell(SheetName, "B10", "C10")
	err = f.MergeCell(SheetName, "D10", "E10")
	err = f.MergeCell(SheetName, "F10", "G10")

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	err = f.SetCellStyle(SheetName, "B11", "G11", style)
	err = f.SetSheetRow(SheetName, "B11", &[]interface{}{"John Doe", "", "#1B800XR", "", "Brand & Marketing"})
	err = f.MergeCell(SheetName, "B11", "C11")
	err = f.MergeCell(SheetName, "D11", "E11")
	err = f.MergeCell(SheetName, "F11", "G11")

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Bold: true}})
	err = f.SetCellStyle(SheetName, "B13", "G13", style)
	err = f.SetSheetRow(SheetName, "B13", &[]interface{}{"Manager", "", "Purpose"})
	err = f.MergeCell(SheetName, "B13", "C13")
	err = f.MergeCell(SheetName, "D13", "E13")

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	err = f.SetCellStyle(SheetName, "B14", "G14", style)
	err = f.SetSheetRow(SheetName, "B14", &[]interface{}{"Jane Doe", "", "Brand Campaign"})
	err = f.MergeCell(SheetName, "B14", "C14")
	err = f.MergeCell(SheetName, "D14", "E14")

	style, err = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "2B4492"},
		Alignment: &excelize.Alignment{Vertical: "center"},
	})
	err = f.SetCellStyle(SheetName, "B17", "G17", style)
	err = f.SetSheetRow(SheetName, "B17", &[]interface{}{"Date", "Category", "Description", "", "Notes", "Amount"})
	err = f.MergeCell(SheetName, "D17", "E17")
	err = f.SetRowHeight(SheetName, 17, 32)

	startRow := 18
	for i := startRow; i < (len(expenseData) + startRow); i++ {
		var fill string
		if i%2 == 0 {
			fill = "F3F3F3"
		} else {
			fill = "FFFFFF"
		}

		style, err = f.NewStyle(&excelize.Style{
			Fill:      excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{fill}},
			Font:      &excelize.Font{Color: "666666"},
			Alignment: &excelize.Alignment{Vertical: "center"},
		})
		err = f.SetCellStyle(SheetName, fmt.Sprintf("B%d", i), fmt.Sprintf("G%d", i), style)
		err = f.SetSheetRow(SheetName, fmt.Sprintf("B%d", i), &expenseData[i-18])
		err = f.SetCellRichText(SheetName, fmt.Sprintf("C%d", i), []excelize.RichTextRun{
			{Text: expenseData[i-18][1].(string), Font: &excelize.Font{Bold: true}},
		})

		err = f.MergeCell(SheetName, fmt.Sprintf("D%d", i), fmt.Sprintf("E%d", i))
		err = f.SetRowHeight(SheetName, i, 18)

	}

	f.SetSheetViewOptions(SheetName, 1, excelize.ShowGridLines(false))

	err = f.SaveAs("expense-report.xlsx")

	err = generateCSV(f, Axis{17, "B"}, Axis{22, "G"})

	if err != nil {
		log.Fatal(err)
	}
}

type Axis struct {
	row int
	col string
}

func generateCSV(f *excelize.File, start, end Axis) error {
	var data [][]string

	for i := start.row; i <= end.row; i++ {
		row := []string{}
		for j := []rune(start.col)[0]; j <= []rune(end.col)[0]; j++ {
			value, err := f.GetCellValue(SheetName, fmt.Sprintf("%s%d", string(j), i), excelize.Options{})
			if err != nil {
				return err
			}
			row = append(row, value)
		}
		data = append(data, row)
	}

	file, err := os.Create("expenses.csv")
	if err != nil {
		return err
	}
	defer f.Close()

	writer := csv.NewWriter(file)
	return writer.WriteAll(data)
}
