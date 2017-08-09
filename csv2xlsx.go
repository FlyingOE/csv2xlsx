package main

import (
	"bufio"
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

var (
	parmCols            string
	parmRows            string
	parmSheet           string
	parmInFile          string
	parmOutFile         string
	parmColSep          rune
	parmDateFormat      string
	parmExcelDateFormat string
	parmNoHeader        bool
	parmSilent          bool
	parmHelp            bool
	parmAbortOnError    bool
	parmShowVersion     bool
	parmAutoFormula     bool
	rowRangeParsed      map[int]string
	colRangeParsed      map[int]string
	workBook            *xlsx.File
	workSheet           *xlsx.Sheet
	rightAligned        *xlsx.Style
	buildTimestamp      string
	versionInfo         string
	tmpStr              string
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

// ParseCommandLine defines and parses command line flags and checks for usage info flags.
// The function exits the program, if the input file does not exist
func parseCommandLine() {
	flag.StringVar(&parmInFile, "infile", "", "full pathname of input file (CSV file)")
	flag.StringVar(&parmOutFile, "outfile", "", "full pathname of output file (.xlsx file)")
	flag.StringVar(&parmDateFormat, "dateformat", "2006-01-02", "format for CSV date cells (default YYYY-MM-DD)")
	flag.StringVar(&parmExcelDateFormat, "exceldateformat", "", "Excel format for date cells (default as in Excel)")
	flag.StringVar(&parmCols, "columns", "", "column range to use (see below)")
	flag.StringVar(&parmRows, "rows", "", "list of line numbers to use (1,2,8 or 1,3-14,28)")
	flag.StringVar(&parmSheet, "sheet", "fromCSV", "tab name of the Excel sheet")
	flag.StringVar(&tmpStr, "colsep", "|", "column separator (default '|') ")
	// not settable with csv reader
	//flag.StringVar(&parmRowSep, "rowsep", "\n", "row separator (default LF) ")
	flag.BoolVar(&parmNoHeader, "noheader", false, "no headers in first line, only data lines (default false)")
	flag.BoolVar(&parmAbortOnError, "abortonerror", false, "abort program on first invalid cell data type")
	flag.BoolVar(&parmSilent, "silent", false, "do not display progress messages")
	flag.BoolVar(&parmAutoFormula, "autoformula", false, "automatically format string starting with = as formulae")
	flag.BoolVar(&parmHelp, "help", false, "display usage information")
	flag.BoolVar(&parmHelp, "h", false, "display usage information")
	flag.BoolVar(&parmHelp, "?", false, "display usage information")
	flag.BoolVar(&parmShowVersion, "version", false, "display version information")
	flag.Parse()

	t, err := strconv.Unquote(`"` + tmpStr + `"`)
	if err != nil {
		fmt.Println("Invalid column separator specified, exiting.")
		os.Exit(1)
	}
	parmColSep, _ = utf8.DecodeRuneInString(t)

	if parmShowVersion {
		fmt.Println("Version ", versionInfo, ", Build timestamp ", buildTimestamp)
		os.Exit(0)
	}

	if parmHelp {
		fmt.Printf("You are running version %s of %s\n\n", versionInfo, filepath.Base(os.Args[0]))
		flag.Usage()
		fmt.Println(`
        Column ranges are a comma-separated list of numbers (e.g. 1,4,8,16), intervals (e.g. 0-4,18-32) or a combination.
        Each comma group can take a type specifiers for the column,
        one of "text", "number", "integer", "currency", date", "standard" or "formula"
        separated from numbers with a colon (e.g. 0:text,3-16:number,17:date)
		`)
		os.Exit(1)
	}
	if _, err := os.Stat(parmInFile); os.IsNotExist(err) {
		fmt.Println("Input file does not exist, exiting.")
		os.Exit(1)
	}
}

// loadInputFile reads the complete input file into a matrix of strings.
// currently there is not need for gigabyte files, but maybe this should be done streaming.
// in addition, we need row and column counts first to set the default ranges later on in the program flow.
func loadInputFile(filename string) (rows [][]string) {
	f, err := os.Open(filename)
	defer f.Close()
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	// use csv reader to read entire file
	r := csv.NewReader(bufio.NewReader(f))
	r.Comma = parmColSep
	r.FieldsPerRecord = -1
	r.LazyQuotes = true
	rows, err = r.ReadAll()
	if err != nil {
		fmt.Println(err)
		os.Exit(2)
	}
	// if we get here, we have file data, so no need for an error value.
	return rows
}

// setRangeInformation uses the input file's row and column count to set the default ranges
// for lines and columns. of course we could leave this out by improving the parser function
// at parseRangeString to allow something like line 34- (instead of 34-999). It's on the list ...
func setRangeInformation(rowCount, colCount int) {
	// now we can set the default ranges for lines and columns
	if parmRows == "" {
		parmRows = fmt.Sprintf("0-%d", rowCount)
	}
	if parmCols == "" {
		parmCols = fmt.Sprintf("0-%d", colCount)
	}
	// will bail out on parse error, see declaration
	rowRangeParsed = parseRangeString(parmRows)
	colRangeParsed = parseRangeString(parmCols)
}

// writeCellContents is basically a boring comparison which data type should be written
// to the spreadsheet cell. if the function encounters invalid values for the data type,
// it outputs an error message and ignores the value
func writeCellContents(cell *xlsx.Cell, colString, colType string, rownum, colnum int) bool {
	success := true
	// only convert to formula if the user specified --autoformula,
	// otherwise use the defined type from column range -- for lazy people :-)
	if parmAutoFormula && []rune(colString)[0] == '=' {
		colType = "formula"
	}
	switch colType {
	case "text":
		cell.SetString(colString)
	case "number","currency":
		floatVal, err := strconv.ParseFloat(colString, 64)
		if err != nil {
			fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid number, value: %s", rownum, colnum, colString))
			success = false
		} else {
			cell.SetStyle(rightAligned)
			if colType == "currency" {
				cell.SetFloatWithFormat(floatVal, "#,##0.00;[red](#,##0.00)")
			} else {
				cell.SetFloat(floatVal)
			}
		}
	case "integer":
		intVal, err := strconv.ParseInt(colString, 10, 64)
		if err != nil {
			fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid integer, value: %s", rownum, colnum, colString))
			success = false
		} else {
			cell.SetStyle(rightAligned)
			cell.SetInt64(intVal)
			cell.NumFmt = "#0"
		}
	case "date":
		dt, err := time.Parse(parmDateFormat, colString)
		if err != nil {
			fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid date, value: %s", rownum, colnum, colString))
			success = false
		} else {
			cell.SetDateTime(dt)
			if parmExcelDateFormat != "" {
				cell.NumFmt = parmExcelDateFormat
			}
		}
	case "formula":
		// colstring =<formula>
		cell.SetFormula(colString[1:])
	default:
		cell.SetValue(colString)
	}
	return success
}

// processDataColumns processes a row from the csv input file and writes a cell for each column
// that should be processed (is in column range, which means it is a key in the colRangeParsed map.
// if the abortOnError option is set, the function exits the program on the first data type error.
func processDataColumns(excelRow *xlsx.Row, rownum int, csvLine []string) {
	if !parmSilent {
		fmt.Println(fmt.Sprintf("Processing csvLine %d (%d cols)", rownum, len(csvLine)))
	}
	for colnum := 0; colnum < len(csvLine); colnum++ {
		colType, processColumn := colRangeParsed[colnum]
		if processColumn {
			cell := excelRow.AddCell()
			if rownum == 0 && !parmNoHeader {
				// special case for the title row
				cell.SetString(csvLine[colnum])
				if colType == "number" || colType == "currency" {
					cell.SetStyle(rightAligned)
				}
			} else {
				// if the user wanted drama (--abortonerror), exit on first error
				ok := writeCellContents(cell, csvLine[colnum], colType, rownum, colnum)
				if !ok && parmAbortOnError {
					os.Exit(3)
				}
			}
		}
	}
}

// the main entry function
func main() {
	// preflight stuff
	parseCommandLine()
	rows := loadInputFile(parmInFile)
	setRangeInformation(len(rows), len(rows[0]))

	// excel stuff, create file, add worksheet, define a right-aligned style
	workBook = xlsx.NewFile()
	workSheet, _ = workBook.AddSheet(parmSheet)
	rightAligned = &xlsx.Style{}
	rightAligned.Alignment = xlsx.Alignment{Horizontal: "right"}

	// loop thru line and column ranges and process data cells
	for rownum := 0; rownum < len(rows); rownum++ {
		_, processLine := rowRangeParsed[rownum]
		if processLine {
			line := rows[rownum]
			excelRow := workSheet.AddRow()
			processDataColumns(excelRow, rownum, line)
		}
	}
	workBook.Save(parmOutFile)
}
