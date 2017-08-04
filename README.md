
## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been crated due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
license below), you may download the binary.

Here are the SHA-256 checksums for the binaries:

	07edbff0609058b31bbbfdce532f0b83919da029555970c12af5fa52b0c2f9d1  csv2xlsx_386.exe
	818041bde85552ea4930152987c478f581614e7768e272af28b2bdd1b4940ed7  csv2xlsx_amd64.exe
	a1c4b4a84e467f878c9ae413e732f697d75fa8ced49c5adea22f9fab251ff3c9  csv2xlsx_linux_386
	955a8d4de854ab0c5fd7c7e676c61f61c3f59712a8da101a1db1c71ebc622bb0  csv2xlsx_linux_amd64
	61ed47ca548ec7773080ff4a059dc2e8a04f345c45b8c0456ef7013dbf9d0047  csv2xlsx_osx


### Usage

You execute the program from a command line shell and prove at least the input file and the output file name.
Please see below for a list of command line options.

### Command line options

```
  -?	display usage information
  -abortonerror
    	abort program on first invalid cell data type
  -colsep string
    	column separator (default '|')  (default "|")
  -columns string
    	column range to use (see below)
  -dateformat string
    	format for CSV date cells (default YYYY-MM-DD) (default "2006-01-02")
  -exceldateformat string
    	Excel format for date cells (default as in Excel)
  -h	display usage information
  -help
    	display usage information
  -infile string
    	full pathname of input file (CSV file)
  -outfile string
    	full pathname of output file (.xlsx file)
  -rows string
    	list of line numbers to use (1,2,8 or 1,3-14,28)
  -rowsep string
    	row separator (default LF)  (default "\n")
  -sheet string
    	tab name of the Excel sheet (default "fromCSV")
  -silent
    	do not display progress messages
  -usetitles
    	use first row as titles (will force string type) (default true)

        Column ranges are a comma-separated list of numbers (e.g. 1,4,8,16), intervals (e.g. 0-4,18-32) or a combination.
        Each comma group can take a type specifiers for the column,
        one of "text", "number", "integer", "currency", date" or "standard",
        separated from numbers with a colon (e.g. 0:text,3-16:number,17:date)
```

### Source

This tool fulfills a special requirement and I will extend its functionality, if need arises. As I found out there are lots 
of people looking for such a tool, I decided to make it publicly available. I am in the process of learning Go and therefore
I am sure there are much better, more Go-idiomatic ways to achieve this functionality. If you have feedback on how to improve
the code or want to contribute, please do not hesitate to do so. I'd really like to improve my GO skills and learn things.
As my spare time for coding is limited to some hours around midnight a week, so please have some patience with my answers.
I am still amazed what you can accomplish within less than 200 lines of code in terms of making my admin part of life easier. :-)

### Changelog

    2017-08-03  0.0.1
                Initial commit. First, ugly version

    2017-08-04  0.1.2
                Refactored code to improve readability, added options
                --abortonerror
                --exceldateformat
                --silent
                Added datatypes integer, currency
                Prints version info on Usage or with --version

### License

This code is licensed under the 2-Clause BSD License:

    Copyright 2017 Armin Hanisch. All rights reserved.

    Redistribution and use in source and binary forms, with or without
    modification, are permitted provided that the following conditions are
    met:

    Redistributions of source code must retain the above copyright notice,
    this list of conditions and the following disclaimer.  Redistributions
    in binary form must reproduce the above copyright notice, this list of
    conditions and the following disclaimer in the documentation and/or
    other materials provided with the distribution.

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
    ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING,
    BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
    FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
    IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR
    ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
    CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
    SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR
    BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
    WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE
    OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN
    IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
