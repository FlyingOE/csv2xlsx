package main

import (
	"github.com/tealeg/xlsx"
	"golang.org/x/text/encoding/charmap"
)

var (
	versionInfo          string = "0.6.0 (2020-03-20)"
	parmCols             string
	parmRows             string
	parmSheet            string
	parmInFile           string
	parmOutFile          string
	parmOutDir           string
	parmFileMask         string
	parmEncoding         string
	parmHeaderLines      int
	parmFontSize         int
	parmFontName         string
	parmColSep           rune
	parmDateFormat       string
	parmExcelDateFormat  string
	parmNoHeader         bool
	parmSilent           bool
	parmHelp             bool
	parmAbortOnError     bool
	parmShowVersion      bool
	parmAutoFormula      bool
	parmIgnoreEmpty      bool
	parmOverwrite        bool
	rowRangeParsed       map[int]string
	colRangeParsed       map[int]string
	workBook             *xlsx.File
	workSheet            *xlsx.Sheet
	rightAligned         *xlsx.Style
	leftAligned          *xlsx.Style
	tmpStr               string
	parmHeaderLabels     []string
	parmAppendToSheet    bool
	parmListEncoders     bool
	parmStartRow         int
	parmNaNValue         string
)

// Possible bailouts
const (
	//SUCCESS             = 0
	SHOW_USAGE          = 1
	INVALID_RANGE       = 2
	INVALID_COLSEP      = 3
	INPUTFILE_NOT_FOUND = 4
	//INPUTFILE_EMPTY     = 5
	//OPEN_CREATE_ERROR   = 6
	//READ_ERROR          = 7
	WRITE_ERROR         = 8
	//ROW_WRITE_ERROR     = 9
	//NO_OVERWRITE        = 10
	OUTPUTDIR_NOT_FOUND = 11
	EXCEL_SAVE_ERROR    = 12
	INVALID_ARGUMENTS   = 13
)

var encoders = map[string]*charmap.Charmap{
	"CODEPAGE037":       charmap.CodePage037,
	"CODEPAGE437":       charmap.CodePage437,
	"CODEPAGE850":       charmap.CodePage850,
	"CODEPAGE852":       charmap.CodePage852,
	"CODEPAGE855":       charmap.CodePage855,
	"CODEPAGE858":       charmap.CodePage858,
	"CODEPAGE860":       charmap.CodePage860,
	"CODEPAGE862":       charmap.CodePage862,
	"CODEPAGE863":       charmap.CodePage863,
	"CODEPAGE865":       charmap.CodePage865,
	"CODEPAGE866":       charmap.CodePage866,
	"CODEPAGE1047":      charmap.CodePage1047,
	"CODEPAGE1140":      charmap.CodePage1140,
	"ISO8859_1":         charmap.ISO8859_1,
	"ISO8859_2":         charmap.ISO8859_2,
	"ISO8859_3":         charmap.ISO8859_3,
	"ISO8859_4":         charmap.ISO8859_4,
	"ISO8859_5":         charmap.ISO8859_5,
	"ISO8859_6":         charmap.ISO8859_6,
	"ISO8859_7":         charmap.ISO8859_7,
	"ISO8859_8":         charmap.ISO8859_8,
	"ISO8859_9":         charmap.ISO8859_9,
	"ISO8859_10":        charmap.ISO8859_10,
	"ISO8859_13":        charmap.ISO8859_13,
	"ISO8859_14":        charmap.ISO8859_14,
	"ISO8859_15":        charmap.ISO8859_15,
	"ISO8859_16":        charmap.ISO8859_16,
	"KOI8R":             charmap.KOI8R,
	"KOI8U":             charmap.KOI8U,
	"MACINTOSH":         charmap.Macintosh,
	"MACINTOSHCYRILLIC": charmap.MacintoshCyrillic,
	"WINDOWS874":        charmap.Windows874,
	"WINDOWS1250":       charmap.Windows1250,
	"WINDOWS1251":       charmap.Windows1251,
	"WINDOWS1252":       charmap.Windows1252,
	"WINDOWS1253":       charmap.Windows1253,
	"WINDOWS1254":       charmap.Windows1254,
	"WINDOWS1255":       charmap.Windows1255,
	"WINDOWS1256":       charmap.Windows1256,
	"WINDOWS1257":       charmap.Windows1257,
	"WINDOWS1258":       charmap.Windows1258,
}
