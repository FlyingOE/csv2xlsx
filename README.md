## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
license below), you may download the binary.

Here are the SHA-256 checksums for the binaries:

    f209b32d6fcc07bb5788789bf38419b80ceace8e37b12fd395adfabce1e861f1  csv2xlsx_386.exe
    76be18318b797742985b4547ddb502688b014bb13f19a1b5d9e886869446e33e  csv2xlsx_amd64.exe
    97fce8d687a73c770dd5463c01ee126bc82fafb1da5fd49dd4751dce06f2cebd  csv2xlsx_linux_386
    a931741383cb68a10bad6a107b0956d5f9e1e23eecb4c11c399bb463307c77bc  csv2xlsx_linux_amd64
    7fbfce2072fedb66fe1a26b1e28399e3ef3b68fc3179778f1c59230fecbda904  csv2xlsx_osx

### Usage

You execute the program from a command line shell and prove at least the input file and the output file name.
Please see below for a list of command line options.

### Command line options

```
  -?	display usage information
  -abortonerror
    	abort program on first invalid cell data type
  -autoformula
        use value starting with an "=" as formula (default False) and do not
        use the column datatype specified
  -colsep string
    	column separator (default '|')  (default "|")
  -columns string
    	column range to use (see below)
  -dateformat string
    	format for CSV date cells (default YYYY-MM-DD) (default "2006-01-02")
  -encoding
      encoding string to use for the CSV file, case-insensitive (defaults to "utf-8")
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
  -sheet string
    	tab name of the Excel sheet (default "fromCSV")
  -silent
    	do not display progress messages
  -noheader
    	do not use the first line as header (default is Yes)

        Column ranges are a comma-separated list of numbers (e.g. 1,4,8,16), intervals (e.g. 0-4,18-32) or a combination.
        Each comma group can take a type specifiers for the column,
        one of "text", "number", "integer", "currency", date", "standard", "formula"
        separated from numbers with a colon (e.g. 0:text,3-16:number,17:date)
```

### Supported encodings

 * Codepage037
 * Codepage437
 * Codepage850
 * Codepage852
 * Codepage855
 * Codepage858
 * Codepage860
 * Codepage862
 * Codepage863
 * Codepage865
 * Codepage866
 * Codepage1047
 * Codepage1140
 * ISO8859_1
 * ISO8859_2
 * ISO8859_3
 * ISO8859_4
 * ISO8859_5
 * ISO8859_6
 * ISO8859_6E
 * ISO8859_6I
 * ISO8859_7
 * ISO8859_8
 * ISO8859_8E
 * ISO8859_8I
 * ISO8859_9
 * ISO8859_10
 * ISO8859_13
 * ISO8859_14
 * ISO8859_15
 * ISO8859_16
 * Koi8r
 * Koi8u
 * Macintosh
 * MacintoshCyrillic
 * Windows874
 * Windows1250
 * Windows1251
 * Windows1252
 * Windows1253
 * Windows1254
 * Windows1255
 * Windows1256
 * Windows1257
 * Windows1258


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

    2017-08-10  0.1.3
                - removed option --usetitles, added --noheader
                - added datatype "formula"
                - option --colsep now handles \t for tab correctly
                - lots of bug fixes

    2017-12-21 0.2
                Added option --encoding

    2017-06-20 0.3
                Added better version of ParseFloat to allow scientific number notation


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
