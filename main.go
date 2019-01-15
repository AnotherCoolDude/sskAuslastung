package main

import (
	"encoding/csv"
	"fmt"
	"os"
)

const (
	filename = "/Users/empfang/Dropbox/test.csv"
)

func main() {
	defer fmt.Println("leaving main...")

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
	fmt.Println(recs)

}
