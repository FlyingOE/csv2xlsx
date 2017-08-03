package main

import (
	"bufio"
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"strconv"
	"strings"
	"time"
)

// parseCommaGroup parses a single comma group (x or x-y),
// optionally followed by :datatype (used only for columns right now)
// It returns a map with row or column index as key and the datatype as value
func parseCommaGroup(grpstr string) (map[int]string, error) {
	var err error
	var startVal int
	var endVal int
	result := make(map[int]string)
	// we need exactly one number or an a-b interval (2 number parts)
	parts := strings.Split(grpstr, "-")
	if len(parts) < 1 || len(parts) > 2 {
		return nil, errors.New(fmt.Sprintf("Invalid range group '%s' found.", grpstr))
	}
	// check for type (currently needed only for columns, will be ignored for lines)
	datatype := "standard"
	// last item may have type spec
	if strings.Index(parts[len(parts)-1], ":") >= 0 {
		datatype = strings.Split(parts[len(parts)-1], ":")[1]
		parts[len(parts)-1] = strings.Split(parts[len(parts)-1], ":")[0]
	}
	// first number
	startVal, err = strconv.Atoi(parts[0])
	if err == nil {
		result[startVal] = datatype
	}
	// interval?
	if len(parts) == 2 {
		endVal, err = strconv.Atoi(parts[1])
		if err == nil {
			for i := startVal + 1; i <= endVal; i++ {
				result[i] = datatype
			}
		}
	}
	return result, err
}

// parseRangeString parses a comma-separated list of range groups.
// It returns a map with row or column index as key and the datatype as value
// As the data source has to be valid, this functions exits the program on parse errors
func parseRangeString(rangeStr string) map[int]string {
	result := make(map[int]string)
	for _, part := range strings.Split(rangeStr, ",") {
		indexlist, err := parseCommaGroup(part)
		if err != nil {
			fmt.Println("Invalid range, exiting.")
			os.Exit(1)
		}
		for key, val := range indexlist {
			result[key] = val
		}
	}
	return result
}

// the main entry function
func main() {
	var (
		parmCols       string
		parmRows       string
		parmSheet      string
		parmInFile     string
		parmOutFile    string
		parmColSep     string
		parmRowSep     string
		parmDateFormat string
		parmUseTitles  bool
		parmSilent     bool
		parmHelp       bool
		err            error
		rowRangeParsed map[int]string
		colRangeParsed map[int]string
		workBook       *xlsx.File
		workSheet      *xlsx.Sheet
	)

	flag.StringVar(&parmInFile, "infile", "", "full pathname of input file (CSV file)")
	flag.StringVar(&parmOutFile, "outfile", "", "full pathname of output file (.xlsx file)")
	flag.StringVar(&parmDateFormat, "dateformat", "2006-01-02", "format for date cells (default YYYY-MM-DD)")
	flag.StringVar(&parmCols, "columns", "", "column range to use (see below)")
	flag.StringVar(&parmRows, "rows", "", "list of line numbers to use (1,2,8 or 1,3-14,28)")
	flag.StringVar(&parmSheet, "sheet", "fromCSV", "tab name of the Excel sheet")
	flag.StringVar(&parmColSep, "colsep", "|", "column separator (default '|') ")
	flag.StringVar(&parmRowSep, "rowsep", "\n", "row separator (default LF) ")
	flag.BoolVar(&parmUseTitles, "usetitles", true, "use first row as titles (will force string type)")
	flag.BoolVar(&parmSilent, "silent", false, "do not display progress messages")
	flag.BoolVar(&parmHelp, "help", false, "display usage information")
	flag.BoolVar(&parmHelp, "h", false, "display usage information")
	flag.BoolVar(&parmHelp, "?", false, "display usage information")
	flag.Parse()

	if parmHelp {
		flag.Usage()
		fmt.Println(`
        Column ranges are a comma-separated list of numbers (e.g. 1,4,8,16), intervals (e.g. 0-4,18-32) or a combination.
        Each comma group can take a type specifiers for the column, one of "text", "number", "date" or "standard",
        separated from numbers with a colon (e.g. 0:text,3-16:number,17:date)
		`)
		os.Exit(1)
	}
	if _, err := os.Stat(parmInFile); os.IsNotExist(err) {
		fmt.Println("Input file does not exist, exiting.")
		os.Exit(1)
	}

	// first read csv file to allow using actual row and col counts as option defaults
	f, err := os.Open(parmInFile)
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	r := csv.NewReader(bufio.NewReader(f))
	r.Comma = []rune(parmColSep)[0]
	r.LazyQuotes = true

	rows, err := r.ReadAll()
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}

	if parmRows == "" {
		parmRows = fmt.Sprintf("0-%d", len(rows))
	}
	if parmCols == "" {
		colCount := len(rows[0])
		parmCols = fmt.Sprintf("0-%d", colCount)
	}

	// will bail out on parse error, see declaration
	rowRangeParsed = parseRangeString(parmRows)
	colRangeParsed = parseRangeString(parmCols)

	workBook = xlsx.NewFile()
	workSheet, _ = workBook.AddSheet(parmSheet)

	for rownum := 0; rownum < len(rows); rownum++ {
		_, ok := rowRangeParsed[rownum]
		if ok {
			line := rows[rownum]
			excelRow := workSheet.AddRow()
			if !parmSilent {
				fmt.Println(fmt.Sprintf("Processing line %d (%d cols)", rownum, len(line)))
			}
			for colnum := 0; colnum < len(line); colnum++ {
				colType, ok := colRangeParsed[colnum]
				if ok {
					cell := excelRow.AddCell()
					if rownum == 0 && parmUseTitles {
						cell.SetString(line[colnum])
					} else {
						switch colType {
						case "text":
							cell.SetString(line[colnum])
						case "number":
							floatVal, err := strconv.ParseFloat(line[colnum], 64)
							if err != nil {
								fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid number, value: %s", rownum, colnum, line[colnum]))
							} else {
								cell.SetFloat(floatVal)
							}
						case "date":
							dt, err := time.Parse(parmDateFormat, line[colnum])
							if err != nil {
								fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid date, value: %s", rownum, colnum, line[colnum]))
							} else {
								cell.SetDateTime(dt)
							}
						default:
							cell.SetValue(line[colnum])
						}
					}
				}
			}
		}
	}
	workBook.Save(parmOutFile)
}
