// Copyright 2011-2015, The xlsx2csv Authors.
// All rights reserved.
// For details, see the LICENSE file.

package main

import (
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx/v3"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
)

func generateCSVFromXLSXFile2(outDir string, excelFileName string, sheetIndex int, csvOpts csvOptSetter) error {
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return err
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	sheet := xlFile.Sheets[sheetIndex]

	f := sheet.Name + ".csv"
	p := filepath.Join(outDir, f)
	w, e := os.Create(p)
	if e != nil {
		log.Fatal(e)
	}
	defer func() {
		if closeErr := w.Close(); closeErr != nil {
			log.Fatal(closeErr)
		}
	}()

	cw := csv.NewWriter(w)
	if csvOpts != nil {
		csvOpts(cw)
	}
	var vals []string
	err = sheet.ForEachRow(func(row *xlsx.Row) error {
		if row != nil {
			vals = vals[:0]
			err := row.ForEachCell(func(cell *xlsx.Cell) error {
				str, err := cell.FormattedValue()
				if err != nil {
					return err
				}
				vals = append(vals, str)
				return nil
			})
			if err != nil {
				return err
			}
		}
		cw.Write(vals)
		return nil
	})
	if err != nil {
		return err
	}
	cw.Flush()
	return cw.Error()
}

func generateCSVFromXLSXFile(w io.Writer, excelFileName string, sheetIndex int, csvOpts csvOptSetter) error {
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return err
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	cw := csv.NewWriter(w)
	if csvOpts != nil {
		csvOpts(cw)
	}
	sheet := xlFile.Sheets[sheetIndex]
	var vals []string
	err = sheet.ForEachRow(func(row *xlsx.Row) error {
		if row != nil {
			vals = vals[:0]
			err := row.ForEachCell(func(cell *xlsx.Cell) error {
				str, err := cell.FormattedValue()
				if err != nil {
					return err
				}
				vals = append(vals, str)
				return nil
			})
			if err != nil {
				return err
			}
		}
		cw.Write(vals)
		return nil
	})
	if err != nil {
		return err
	}
	cw.Flush()
	return cw.Error()
}

type csvOptSetter func(*csv.Writer)

func main() {
	var (
		//outDir     = flag.String("o", ".", "dir to output to")
		sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")
		delimiter  = flag.String("d", "\t", "Delimiter to use between fields")
	)
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, "Usage: %s [flags] <dir-to-be-read> <dir-to-output-to>\n", os.Args[0])
		flag.PrintDefaults()
	}

	flag.Parse()
	if flag.NArg() != 2 {
		flag.Usage()
		os.Exit(1)
	}

	filepath.Walk(flag.Arg(0), func(path string, info os.FileInfo, err error) error {
		//忽略目录
		if info.IsDir() {
			return nil
		}
		//是不是xlsx
		if strings.Compare(filepath.Ext(path), ".xlsx") != 0 {
			return nil
		}

		outDir := flag.Arg(1)
		os.MkdirAll(outDir, 0755)

		if err := generateCSVFromXLSXFile2(outDir, path, *sheetIndex, func(cw *csv.Writer) {
			cw.Comma = ([]rune(*delimiter))[0]
		}); err != nil {
			//log.Fatal(err)
			log.Println(err)
		}
		return nil
	})
}
