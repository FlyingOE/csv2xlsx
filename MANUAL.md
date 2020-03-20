
# Usage information

## List of command line options

      -?	display usage information
      -abortonerror
            abort program on first invalid cell data type
      -append
            append data rows to specified sheet instead of overwriting sheet
      -autoformula
            automatically format string starting with = as formulae
      -colsep string
            column separator (default '|')  (default "|")
      -columns string
            column range to use (see below)
      -dateformat string
            format for CSV date cells (default YYYY-MM-DD) (default "2006-01-02")
      -encoding string
            character encoding (default "utf-8")
      -exceldateformat string
            Excel format for date cells (default as in Excel)
      -filemask string
            file mask for bulk processing (overwrites -infile/-outfile)
      -fontname string
            set the font name to use (default "Arial")
      -fontsize int
            set the default font size to use (default 12)
      -h	display usage information
      -headerlabels string
            comma-separated list of header labels (enclose in quotes to be safe)
      -headerlines int
            set the number of header lines (use 0 for no header) (default 1)
      -help
            display usage information
      -ignoreempty
            do not display warnings for empty cells (default true)
      -infile string
            full pathname of input file (CSV file)
      -listencodings
            display a list of supported encodings and exit
      -nanvalue string
            value to be used for failed number conversions or missing numbers
      -noheader
            DEPRECATED (use headerlines) no header, only data lines
      -outdir string
            target directory for the .xlsx file (not to be used with outfile)
      -outfile string
            full pathname of output file (.xlsx file)
      -overwrite
            overwrite existing output file (default false)
      -rows string
            list of line numbers to use (1,2,8 or 1,3-14,28)
      -sheet string
            tab name of the Excel sheet (default "fromCSV")
      -silent
            do not display progress messages
      -startrow int
            start at row N in CSV file (this value is 1-based!) (default 1)
      -version
            display version information
    
        Column ranges are a comma-separated list of numbers (e.g. 1,4,8,16), intervals (e.g. 0-4,18-32) or a combination.
        Each comma group can have type specifier for the columns, separated with a colon (e.g. 0:text,3-16:number,17:date)
        Type is one of: text|number|interger|currency|date|standard|percent|formula|format
		Type "format" may be used together with a format string: format="FMTSTR", e.g. 2:format="0000.0"