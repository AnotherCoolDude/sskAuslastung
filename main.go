package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"github.com/alecthomas/kingpin"
	"github.com/buger/goterm"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var (
	app       = kingpin.New("Auslastung", "füllt automatisch die Excel-Datei Auslastung mit dem Proad-Export aus")
	period    = app.Flag("Zeitraum", "Zeitraum der Proad-Datei. Falls nicht angegeben, wir der Name der Proad-Datei verwendet").Short('z').String()
	excelPath = app.Flag("excelPath", "Pfad zu der Excel-Datei").Required().Short('e').String()
	proadPath = app.Flag("proadPath", "Pfad zu der Proad-Datei").Required().Short('p').String()
	destPath  = app.Flag("destPath", "ein anderer Speicherort für die Excel-Datei").Short('d').String()
	parseOnly = app.Flag("csv_only", "nur die Proad-Datei verarbeiten und anzeigen").Short('o').Bool()
	dontSafe  = app.Flag("dont_safe", "die Änderungen werden nicht gespeichert").Short('s').Bool()

	freelancer = []string{
		"Tina Botz",
	}
)

const (
	jobNrOvertime = "SEIN-0001-0137"
	jobNrNoWork   = "SEIN-0001-0113"
	jobNrSick     = "SEIN-0001-0015"
	jobNrVacation = "SEIN-0001-0012"

	overtime recordType = 8
	noWork   recordType = 3
	sick     recordType = 6
	vacation recordType = 5
	intern   recordType = 4
	customer recordType = 1
	pitch    recordType = 2
)

func main() {
	defer fmt.Println("\nleaving main...")

	kingpin.MustParse(app.Parse(os.Args[1:]))
	ePath := *excelPath
	pPath := *proadPath
	dPath := *destPath
	p := *period

	recs := recordCollection{}
	switch filepath.Ext(ePath) {
	case ".xlsx":
		recs = parseRecordsXLSX(pPath)
	case ".csv":
		recs = parseRecordsCSV(pPath)
	default:
		fmt.Println("only .xlsx and .csv files are supported")
		os.Exit(0)
	}

	assignRecords(recs)
	recs.list()
	if *parseOnly {
		os.Exit(0)
	}

	xlsx, err := excelize.OpenFile(ePath)
	if err != nil {
		fmt.Println(err)
	}
	if p == "" {
		p = strings.TrimSuffix(filepath.Base(pPath), filepath.Ext(pPath))
	}
	recs.addToExcel(xlsx, p)

	if !*dontSafe {
		if dPath != "" {
			saveExcel(xlsx, dPath)
		} else {
			saveExcel(xlsx, ePath)
		}
	}
}

//csv

type jobrecord struct {
	shortName   string
	name        string
	activity    string
	jobDesc     string
	jobNr       string
	workingTime float32
	registered  bool
	recType     recordType
}

type recordType int

func parseRecordsXLSX(filePath string) recordCollection {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	data := file.GetRows(file.GetSheetName(file.GetActiveSheetIndex()))
	return parseRecords(data)
}

func parseRecordsCSV(filePath string) recordCollection {
	file, err := os.Open(filePath)

	if err != nil {
		fmt.Printf("error opening file %s: %s", filePath, err)
	} else {
		stat, _ := file.Stat()
		fmt.Printf("using file: %v, size: %d\n", stat.Name(), stat.Size())
	}
	defer file.Close()

	r := csv.NewReader(file)
	r.Comma = ','
	r.Comment = '#'

	recs, err := r.ReadAll()
	if err != nil {
		fmt.Println(err)
	}
	return parseRecords(recs)
}

func parseRecords(data [][]string) recordCollection {
	jobrecords := []jobrecord{}
	for _, row := range data {
		wt, _ := strconv.ParseFloat(row[8], 32)
		newRecord := jobrecord{
			shortName:   row[0],
			name:        row[1],
			activity:    row[3],
			jobDesc:     row[6],
			jobNr:       row[7],
			workingTime: float32(wt),
			registered:  false,
		}
		jobrecords = append(jobrecords, newRecord)
	}
	return jobrecords[1:]
}

func assignRecords(recs []jobrecord) {
	for i, rec := range recs {
		//ignore Freelancer
		if contains(freelancer, rec.name) {
			continue
		}

		if rec.jobNr == jobNrVacation {
			recs[i].registered = true
			recs[i].recType = vacation
		}
		if rec.jobNr == jobNrSick {
			recs[i].registered = true
			recs[i].recType = sick
		}
		if rec.jobNr == jobNrNoWork {
			recs[i].registered = true
			recs[i].recType = noWork
		}
		if rec.jobNr == jobNrOvertime {
			recs[i].registered = true
			recs[i].recType = overtime
		}
		if caseInsensitiveContains(recs[i].jobDesc, "pitch") {
			recs[i].registered = true
			recs[i].recType = pitch
		}
	}

	for i := range recs {
		if strings.Contains(recs[i].jobNr, "SEIN") && !recs[i].registered {
			recs[i].registered = true
			recs[i].recType = intern
		} else if !recs[i].registered {
			recs[i].registered = true
			recs[i].recType = customer
		}
	}

	fmt.Println("not registered records:")
	for i := range recs {
		if !recs[i].registered {
			fmt.Println(recs[i])
		}
	}
	fmt.Println()
}

func createCSV([][]string) {

}

//excerlize

type recordCollection []jobrecord

func (collection recordCollection) addToExcel(xlsx *excelize.File, period string) {
	table := goterm.NewTable(0, 10, 6, ' ', 0)
	for _, rt := range recordTypes() {
		sheetName := xlsx.GetSheetMap()[int(rt)]
		if sheetName == "" {
			fmt.Printf("\nworksheet %v was not found in file %s\n", rt, xlsx.Path)
			continue
		}
		col, row := getNextFreeCell(xlsx, sheetName)
		if col == "" {
			fmt.Println("next free cell couldn't be determed")
		}
		coordsZeitraum := fmt.Sprintf("%s%s", col, strconv.Itoa(row))
		xlsx.SetCellStr(sheetName, coordsZeitraum, period)

		fmt.Fprintf(table, "%s\t\t\n", sheetName)
		fmt.Fprintf(table, "Coords\tName\tValue\n")
		for _, rec := range collection {
			if rec.recType != rt {
				continue
			}
			dc, name, cv := setValueForEmployee(xlsx, sheetName, col, rec.name, rec.workingTime)
			fmt.Fprintf(table, "%s\t%s\t%f\n", dc, name, cv)
		}
		fmt.Fprintf(table, "\n")
		fmt.Println(table.String())
	}

}

func saveExcel(file *excelize.File, path string) {
	err := file.SaveAs(path)
	if err != nil {
		fmt.Println(err)
	}
}

func coords(coord string) (column string, row int) {
	reg := regexp.MustCompile("[0-9]+|[A-Z]+")
	result := reg.FindAllString(coord, 2)
	n, _ := strconv.Atoi(result[1])
	return result[0], n
}

func getNextFreeCell(file *excelize.File, sheetName string) (column string, row int) {
	rows := file.GetRows(sheetName)
	coordsZeitraum := file.SearchSheet(sheetName, "Zeitraum")
	_, n := coords(coordsZeitraum[0])
	for i, value := range rows[n-1] {
		if value == "" {
			return excelize.ToAlphaString(i), n
		}
	}
	return "", -1
}

func setValueForEmployee(file *excelize.File, sheetname, column, employeename string, value float32) (destcoord, name string, currentValue float32) {
	names := strings.Split(employeename, " ")
	employeeCoords := file.SearchSheet(sheetname, fmt.Sprintf("(%s).*(%s)|(%s).*(%s)", names[0], names[1], names[1], names[0]), true)
	if len(employeeCoords) != 1 {
		fmt.Printf("\n%s either not found or exists more than once \n", employeename)
		fmt.Println(employeeCoords)
		return "n/a", employeename, 0.0
	}
	_, employeeNumber := coords(employeeCoords[0])
	destCoords := fmt.Sprintf("%s%s", column, strconv.Itoa(employeeNumber))
	cellValueString := file.GetCellValue(sheetname, destCoords)
	cellValue, err := strconv.ParseFloat(cellValueString, 32)
	if err != nil {
		cellValue = 0.0
	}
	file.SetCellValue(sheetname, destCoords, value+float32(cellValue))
	return destCoords, employeename, value + float32(cellValue)
}

//helper

func caseInsensitiveContains(s, substr string) bool {
	s, substr = strings.ToUpper(s), strings.ToUpper(substr)
	return strings.Contains(s, substr)
}

func contains(slice []string, str string) bool {
	for _, s := range slice {
		if s == str {
			return true
		}
	}
	return false
}

func recordTypes() []recordType {
	return []recordType{
		overtime,
		noWork,
		sick,
		vacation,
		intern,
		customer,
		pitch,
	}
}

func (rectype recordType) toString() string {
	switch rectype {
	case overtime:
		return "overtime"
	case noWork:
		return "noWork"
	case sick:
		return "sick"
	case vacation:
		return "vacation"
	case intern:
		return "intern"
	case customer:
		return "customer"
	case pitch:
		return "pitch"
	default:
		return ""
	}
}

func (collection recordCollection) list() {
	for _, rectype := range recordTypes() {
		fmt.Println()
		fmt.Println(rectype.toString())
		table := goterm.NewTable(0, 10, 6, ' ', 0)
		fmt.Fprintf(table, "Employee\tHours\n")
		for _, rec := range collection {
			if rec.recType != rectype {
				continue
			}
			fmt.Fprintf(table, "%s\t%f\n", rec.shortName, rec.workingTime)
		}
		fmt.Fprintf(table, "\n")
		fmt.Print(table.String())
	}
}
