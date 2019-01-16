package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const (
	csvPath       = "/Users/christianhovenbitzer/Dropbox/test.csv"
	excelPath     = "/Users/christianhovenbitzer/Dropbox/übersicht Auslastung  Jan_Okt2018.xlsx"
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
	defer fmt.Println("leaving main...")
	recs := parseRecords(csvPath)
	assignRecords(recs)
	fmt.Println("opening xlsx file:")
	xlsx, err := excelize.OpenFile(excelPath)
	if err != nil {
		fmt.Println(err)
	}
	//zeitraum := "zeitraum"
	sheetName := xlsx.GetSheetMap()[int(vacation)]
	l, n := getNextFreeCell(xlsx, sheetName)
	coordsZeitraum := fmt.Sprintf("%s%s", l, strconv.Itoa(n))
	fmt.Println(coordsZeitraum)
	//xlsx.SetCellStr(sheetName, coordsZeitraum, zeitraum)
	for _, rec := range recs {
		if rec.recType != vacation {
			continue
		}
		setIntForEmployee(xlsx, sheetName, rec.name)
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

func parseRecords(filePath string) []jobrecord {
	file, err := os.Open(filePath)

	if err != nil {
		fmt.Printf("error opening file %s: %s", filePath, err)
	} else {
		stat, _ := file.Stat()
		fmt.Printf("using file: %v, size: %d\n\n\n", stat.Name(), stat.Size())
	}
	defer file.Close()

	r := csv.NewReader(file)
	r.Comma = ','
	r.Comment = '#'

	recs, err := r.ReadAll()
	if err != nil {
		fmt.Println(err)
	}
	jobrecords := []jobrecord{}
	for _, rec := range recs {
		wt, _ := strconv.ParseFloat(rec[8], 32)
		newRecord := jobrecord{
			shortName:   rec[0],
			name:        rec[1],
			activity:    rec[3],
			jobDesc:     rec[6],
			jobNr:       rec[7],
			workingTime: float32(wt),
			registered:  false,
		}
		jobrecords = append(jobrecords, newRecord)
	}
	return jobrecords[1:]
}

func assignRecords(recs []jobrecord) {
	for i, rec := range recs {
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

//excerlize

func coords(coord string) (letter string, number int) {
	runes := []rune(coord)
	return string(runes[0]), int(runes[1] - '0')
}

func getNextFreeCell(file *excelize.File, sheetName string) (letter string, number int) {
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

func setIntForEmployee(file *excelize.File, sheetname, employeename string) {
	names := strings.Split(employeename, " ")
	employeeCoords := file.SearchSheet(sheetname, fmt.Sprintf("(?i)(\\W|^)(%s|%s)(\\W|$)", names[0], names[1]), true)
	if len(employeeCoords) != 1 {
		fmt.Printf("\n%s either not found or exists more than once \n", employeename)
		fmt.Println(employeeCoords)
		return
	}
	_, employeeNumber := coords(employeeCoords[0])
	//newCellLetter, _ := getNextFreeCell(file, sheetname)
	value := file.GetCellValue(sheetname, fmt.Sprintf("AR" /*newCellLetter*/, strconv.Itoa(employeeNumber)))
	fmt.Println(value)
}

//helper

func caseInsensitiveContains(s, substr string) bool {
	s, substr = strings.ToUpper(s), strings.ToUpper(substr)
	return strings.Contains(s, substr)
}

func sliceInfo(name string, slice []jobrecord) {
	fmt.Printf("\n%s:\n", name)
	for _, r := range slice {
		fmt.Println(r)
	}
	fmt.Println(len(slice))
	fmt.Println()
}

func formatName(name string) string {
	n := strings.Split(name, " ")
	return fmt.Sprintf("%s, %s", strings.TrimSpace(n[1]), strings.TrimSpace(n[0]))
}
