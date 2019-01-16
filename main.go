package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"strconv"
	"strings"
)

const (
	filename      = "/Users/christianhovenbitzer/Dropbox/test.csv"
	jobNrOvertime = "SEIN-0001-0137"
	jobNrNoWork   = "SEIN-0001-0113"
	jobNrSick     = "SEIN-0001-0015"
	jobNrVacation = "SEIN-0001-0012"

	overtime recordType = 0
	noWork   recordType = 1
	sick     recordType = 2
	vacation recordType = 3
	intern   recordType = 4
	customer recordType = 5
	pitch    recordType = 6
)

func main() {
	defer fmt.Println("leaving main...")
	recs := parseRecords(filename)

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
	file, err := os.Open(filename)

	if err != nil {
		fmt.Printf("error opening file %s: %s", filename, err)
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
