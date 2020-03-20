package main

import (
	"bufio"
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/tealeg/xlsx"
)

// listEncoders is a helper function to display the list
// of supported encodings on standard output
func listEncoders() {
	names := make([]string, 0, len(encoders))
	for name := range encoders {
		names = append(names, name)
	}
	sort.Strings(names) //sort by key
	for enc := range names {
		fmt.Println(names[enc])
	}
}

// ParseFloat ist an advanced ParseFloat for golang, support scientific notation, comma separated number
// from yyscamper at https://gist.github.com/yyscamper/5657c360fadd6701580f3c0bcca9f63a
func ParseFloat(str string) (float64, error) {
	val, err := strconv.ParseFloat(str, 64)
	if err == nil {
		return val, nil
	}
	//Some number may be seperated by comma, for example, 23,120,123, so remove the comma firstly
	str = strings.Replace(str, ",", "", -1)
	//Some number is specifed in scientific notation
	pos := strings.IndexAny(str, "eE")
	if pos < 0 {
		return strconv.ParseFloat(str, 64)
	}
	var baseVal float64
	var expVal int64
	baseStr := str[0:pos]
	baseVal, err = strconv.ParseFloat(baseStr, 64)
	if err != nil {
		return 0, err
	}
	expStr := str[(pos + 1):]
	expVal, err = strconv.ParseInt(expStr, 10, 64)
	if err != nil {
		return 0, err
	}
	return baseVal * math.Pow10(int(expVal)), nil
}

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
		return nil, fmt.Errorf("Invalid range group '%s' found.", grpstr)
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
			os.Exit(INVALID_RANGE)
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
	var headerString = ""
	cmdlineFlags := flag.NewFlagSet(os.Args[0], flag.PanicOnError)
	cmdlineFlags.StringVar(&parmInFile, "infile", "", "full pathname of input file (CSV file)")
	cmdlineFlags.StringVar(&parmOutFile, "outfile", "", "full pathname of output file (.xlsx file)")
	cmdlineFlags.StringVar(&parmFileMask, "filemask", "", "file mask for bulk processing (overwrites -infile/-outfile)")
	cmdlineFlags.StringVar(&parmOutDir, "outdir", "", "target directory for the .xlsx file (not to be used with outfile)")
	cmdlineFlags.StringVar(&parmDateFormat, "dateformat", "2006-01-02", "format for CSV date cells (default YYYY-MM-DD)")
	cmdlineFlags.StringVar(&parmExcelDateFormat, "exceldateformat", "", "Excel format for date cells (default as in Excel)")
	cmdlineFlags.StringVar(&parmCols, "columns", "", "column range to use (see below)")
	cmdlineFlags.StringVar(&parmRows, "rows", "", "list of line numbers to use (1,2,8 or 1,3-14,28)")
	cmdlineFlags.StringVar(&parmSheet, "sheet", "fromCSV", "tab name of the Excel sheet")
	cmdlineFlags.StringVar(&tmpStr, "colsep", "|", "column separator (default '|') ")
	cmdlineFlags.StringVar(&parmEncoding, "encoding", "utf-8", "character encoding")
	cmdlineFlags.StringVar(&parmFontName, "fontname", "Arial", "set the font name to use")
	cmdlineFlags.StringVar(&headerString, "headerlabels", "", "comma-separated list of header labels (enclose in quotes to be safe)")
	cmdlineFlags.IntVar(&parmFontSize, "fontsize", 12, "set the default font size to use")
	cmdlineFlags.IntVar(&parmHeaderLines, "headerlines", 1, "set the number of header lines (use 0 for no header)")
	cmdlineFlags.BoolVar(&parmNoHeader, "noheader", false, "DEPRECATED (use headerlines) no header, only data lines")
	cmdlineFlags.BoolVar(&parmAbortOnError, "abortonerror", false, "abort program on first invalid cell data type")
	cmdlineFlags.BoolVar(&parmSilent, "silent", false, "do not display progress messages")
	cmdlineFlags.BoolVar(&parmAutoFormula, "autoformula", false, "automatically format string starting with = as formulae")
	cmdlineFlags.BoolVar(&parmHelp, "help", false, "display usage information")
	cmdlineFlags.BoolVar(&parmHelp, "h", false, "display usage information")
	cmdlineFlags.BoolVar(&parmHelp, "?", false, "display usage information")
	cmdlineFlags.BoolVar(&parmShowVersion, "version", false, "display version information")
	cmdlineFlags.BoolVar(&parmIgnoreEmpty, "ignoreempty", true, "do not display warnings for empty cells")
	cmdlineFlags.BoolVar(&parmOverwrite, "overwrite", false, "overwrite existing output file (default false)")
	cmdlineFlags.BoolVar(&parmAppendToSheet, "append", false, "append data rows to specified sheet instead of overwriting sheet")
	cmdlineFlags.BoolVar(&parmListEncoders, "listencodings", false, "display a list of supported encodings and exit")
	cmdlineFlags.IntVar(&parmStartRow, "startrow", 1, "start at row N in CSV file (this value is 1-based!)")
	cmdlineFlags.StringVar(&parmNaNValue, "nanvalue", "", "value to be used for failed number conversions or missing numbers")
	err := cmdlineFlags.Parse(os.Args[1:])

	if parmHelp {
		fmt.Printf("You are running version %s of %s\n\n", versionInfo, os.Args[0])
		cmdlineFlags.Usage()
		fmt.Println(`
        Column ranges are a comma-separated list of numbers (e.g. 1,4,8,16), intervals (e.g. 0-4,18-32) or a combination.
        Each comma group can have type specifier for the columns, separated with a colon (e.g. 0:text,3-16:number,17:date)
        Type is one of: text|number|interger|currency|date|standard|percent|formula|format
		Type "format" may be used together with a format string: format="FMTSTR", e.g. 2:format="0000.0"
		`)
		os.Exit(SHOW_USAGE)
	}

	t, err := strconv.Unquote(`"` + tmpStr + `"`)
	if err != nil {
		fmt.Println("Invalid column separator specified, exiting.")
		os.Exit(INVALID_COLSEP)
	}
	parmColSep, _ = utf8.DecodeRuneInString(t)

	if parmShowVersion {
		fmt.Println("Version ", versionInfo)
		os.Exit(SHOW_USAGE)
	}

	r := strings.NewReplacer("YYYY", "2006",
		"MM", "01",
		"DD", "02",
		"HH", "15",
		"MI", "04",
		"SS", "05",
		"ZN", "-7000",
		"ZC", "-07:00")

	// Replace all pairs.
	parmDateFormat = r.Replace(parmDateFormat)

	// do we have user defined header labels?
	parmHeaderLabels = []string{}
	if headerString != "" {
		parmHeaderLabels = strings.Split(headerString, ",")
		for i := range parmHeaderLabels {
			parmHeaderLabels[i] = strings.TrimSpace(parmHeaderLabels[i])
		}
	}

	if parmListEncoders {
		listEncoders()
		os.Exit(SHOW_USAGE)
	}

	if parmOutFile != "" && parmOutDir != "" {
		fmt.Println("Cannot use -outfile and -outdir together (-outdir to be used with -filemask), exiting.")
		os.Exit(INVALID_ARGUMENTS)
	}

	if parmFileMask == "" {
		if _, err := os.Stat(parmInFile); os.IsNotExist(err) {
			fmt.Println("Input file does not exist, exiting.")
			os.Exit(INPUTFILE_NOT_FOUND)
		}
	}
}

// loadInputFile reads the complete input file into a matrix of strings.
// currently there is not need for gigabyte files, but maybe this should be done streaming.
// in addition, we need row and column counts first to set the default ranges later on in the program flow.
func loadInputFile(filename string) (rows [][]string, err error) {
	var rdr io.Reader

	// check if file exists
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		return nil, errors.New(fmt.Sprintf("Input file %s does not exist", filename))
	}

	// open input file
	f, err := os.Open(filename)
	if err != nil {
		return nil, errors.New(fmt.Sprintf("Error opening input file %s", filename))
	}

	// check encoding for input file
	encname := strings.ToUpper(parmEncoding)
	if encname == "UTF8" || encname == "UTF-8" {
		rdr = bufio.NewReader(f)
	} else {
		if enc, ok := encoders[encname]; ok {
			rdr = enc.NewDecoder().Reader(f)
		} else {
			fmt.Println(fmt.Sprintf("Specified encoding \"%s\" not found, defaulting to UTF-8\n", parmEncoding))
			rdr = bufio.NewReader(f)
		}
	}

	// read file data
	r := csv.NewReader(rdr)
	r.Comma = parmColSep
	r.FieldsPerRecord = -1
	r.LazyQuotes = true
	rows, err = r.ReadAll()
	if err != nil {
		msg := fmt.Sprintf("Error reading CSV file %s", filename)
		closeErr := f.Close()
		if closeErr != nil {
			msg = fmt.Sprintf("Error closing file %s", filename) + msg
		}
		return nil, errors.New(msg)
	}
	closeErr := f.Close()
	if closeErr != nil {
		msg := fmt.Sprintf("Error closing file %s", filename)
		return nil, errors.New(msg)
	}

	return rows, nil
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
	theStyle := leftAligned // let's assume left-aligned
	// check for content to process and
	// process the "Ignore Warnings On Empty Cells" flag
	if colString == "" {
		if !parmIgnoreEmpty {
			fmt.Println(fmt.Sprintf("Warning: Cell (%d, %d) is empty.", rownum, colnum))
		}
		return true
	}
	// only convert to formula if the user specified --autoformula,
	// otherwise use the defined type from column range -- for lazy people :-)
	if parmAutoFormula && strings.HasPrefix(colString, "=") {
		colType = "formula"
	}
	// special treatment of "format" column type
	fmtstring := ""
	if strings.HasPrefix(colType, "format") {
		parts := strings.Split(colType, "=")
		if len(parts) > 1 {
			colType = parts[0]
			fmtstring = parts[1]
		}
	}
	// type dependent write
	//fmt.Println("==>", colType)
	switch colType {
	case "format":
		floatVal, err := ParseFloat(colString)
		if err != nil {
			fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid number, value: %s", rownum, colnum, colString))
			success = false
		} else {
			theStyle = rightAligned
			cell.SetFloatWithFormat(floatVal, fmtstring)
			// cell.SetFloatWithFormat(floatVal, fmtstring)
		}
	case "text":
		cell.SetString(colString)
	case "number", "currency":
		floatVal, err := ParseFloat(colString)
		if err != nil {
			if parmNaNValue != "" {
				cell.SetString(parmNaNValue)
				success = true
			} else {
				fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid number, value: %s", rownum, colnum, colString))
				success = false
			}
		} else {
			theStyle = rightAligned
			if colType == "currency" {
				cell.SetFloatWithFormat(floatVal, "#,##0.00;[red](#,##0.00)")
			} else {
				cell.SetFloatWithFormat(floatVal, "0#.###")
			}
		}
	case "integer":
		intVal, err := strconv.ParseInt(colString, 10, 64)
		if err != nil {
			fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid integer, value: %s", rownum, colnum, colString))
			success = false
		} else {
			theStyle = rightAligned
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
	case "percent":
		// thanks to Felipe Augusto da Silva for the improvement to use "percent"
		floatVal, err := ParseFloat(colString)
		if err != nil {
			fmt.Println(fmt.Sprintf("Cell (%d,%d) is not a valid number, value: %s", rownum, colnum, colString))
			success = false
		} else {
			theStyle = rightAligned
			cell.SetFloatWithFormat(floatVal, "0.00%")
		}
	default:
		cell.SetValue(colString)
		_, err := ParseFloat(colString)
		if err == nil {
			theStyle = rightAligned
		}
	}
	cell.SetStyle(theStyle)
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
			isHeader := (parmHeaderLines > 0) || !parmNoHeader
			if isHeader && (rownum < parmHeaderLines) {
				// special case for the title row
				if len(parmHeaderLabels) > 0 && len(parmHeaderLabels) > colnum {
					cell.SetString(parmHeaderLabels[colnum])
				} else {
					cell.SetString(csvLine[colnum])
				}
				cell.SetStyle(leftAligned)
			} else {
				// if the user wanted drama (--abortonerror), exit on first error
				ok := writeCellContents(cell, csvLine[colnum], colType, rownum, colnum)
				if !ok && parmAbortOnError {
					os.Exit(WRITE_ERROR)
				}
			}
		}
	}
}

// getInputFiles retrieves a list of input files for a given filespec
// returns a slice of strings or aborts the program on error
func getInputFiles(inFileSpec string) []string {
	files, err := filepath.Glob(inFileSpec)
	if err != nil {
		fmt.Println(err)
		os.Exit(INPUTFILE_NOT_FOUND)
	}
	return files
}

// buildOutputName generates the .xlsx file name for a given input file
// if the user specified the -outdir option, use this directory as target path
// return a string with the target file name
func buildOutputName(infile string) string {
	outfile := strings.TrimSuffix(infile, filepath.Ext(infile)) + ".xlsx"
	if parmOutFile != "" {
		outfile = parmOutFile
	}
	if parmOutDir != "" {
		if _, err := os.Stat(parmOutDir); err == nil {
			outfile = filepath.Join(parmOutDir, filepath.Base(outfile))
		} else {
			fmt.Println(fmt.Sprintf("Output directory %q does not exist, exiting.", parmOutDir))
			os.Exit(OUTPUTDIR_NOT_FOUND)
		}
	}
	return outfile
}

// openOrCreateFile checks if the specified filename exists and
// tries to read the file subsequently. If the file does not exist,
// a new file instance is created
func openOrCreateFile(filename string) (*xlsx.File, error) {
	var err error
	var f *xlsx.File
	if _, err = os.Stat(filename); os.IsNotExist(err) {
		f = xlsx.NewFile()
		err = nil
	} else {
		f, err = xlsx.OpenFile(filename)
	}
	return f, err
}

// getWorkSheet retrieves the specified sheet from the workbook
// if the sheet does not exist, it is appended to the file.
// Returns a pointer to the sheet
func getWorkSheet(sheetName string, workBook *xlsx.File, appendSheet bool) *xlsx.Sheet {
	var sh *xlsx.Sheet
	var ok bool
	if sh, ok = workBook.Sheet[sheetName]; !ok {
		sh, _ = workBook.AddSheet(sheetName)
	} else {
		if !appendSheet {
			// make a new sheet
			sh, _ = xlsx.NewSheet(sheetName)
		}
	}
	return sh
}

// convertFile does the conversion from CSV to Excel .xslx
func convertFile(infile, outfile string) bool {
	if _, err := os.Stat(outfile); err == nil {
		if !parmOverwrite {
			fmt.Println(fmt.Sprintf("Output file %s exists, skipping (use --overwrite?)", outfile))
			return false
		}
	}

	rows, err := loadInputFile(infile)
	if err != nil {
		fmt.Println(err)
		return false
	}
	setRangeInformation(len(rows), len(rows[0]))

	// excel stuff, create file, add worksheet, define a right-aligned style
	xlsx.SetDefaultFont(parmFontSize, parmFontName)
	rightAligned = &xlsx.Style{}
	rightAligned.Alignment = xlsx.Alignment{Horizontal: "right"}
	rightAligned.Font.Name = parmFontName
	rightAligned.Font.Size = parmFontSize
	leftAligned = &xlsx.Style{}
	leftAligned.Alignment = xlsx.Alignment{Horizontal: "left"}
	leftAligned.Font.Name = parmFontName
	leftAligned.Font.Size = parmFontSize
	workBook, err = openOrCreateFile(outfile)
	if err != nil {
		fmt.Println(err)
		return false
	}
	workSheet = getWorkSheet(parmSheet, workBook, parmAppendToSheet)

	// loop thru line and column ranges and process data cells
	for rownum := 0; rownum < len(rows); rownum++ {
		_, processLine := rowRangeParsed[rownum]
		if processLine {
			line := rows[rownum]
			excelRow := workSheet.AddRow()
			processDataColumns(excelRow, rownum, line)
		}
	}
	err = workBook.Save(outfile)
	if err != nil {
		fmt.Println(err)
		os.Exit(EXCEL_SAVE_ERROR)
	}

	return true
}

func main() {
	var fileList []string
	parseCommandLine()

	// either glob the file mask or retrieve the single input file
	// this way we can just iterate over the slice (and maybe later
	// add an option for a specified list of files)
	if parmFileMask != "" {
		if !parmSilent {
			fmt.Println("Found filemask parameter, running in bulk mode")
		}
		fileList = getInputFiles(parmFileMask)
	} else {
		fileList = getInputFiles(parmInFile)
	}

	// iterate over list of files to process and convert them
	for _, infile := range fileList {
		outfile := buildOutputName(infile)
		if !parmSilent {
			fmt.Println("Converting", infile, "=>", outfile)
		}
		ok := convertFile(infile, outfile)
		if !ok {
			fmt.Println(fmt.Sprintf("Could not convert input file %s", infile))
		}
	}
}
