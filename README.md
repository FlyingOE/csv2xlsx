
## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been crated due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
license below), you may download the binary.

Here are the SHA-256 checksums for the binaries:

    724ffe87ca8b81173faf3219015fb3e529ef357399b5827bafe564b8c8d87970  csv2xlsx_386.exe
    7068b39ac35b2a419fbde39871253170c4666a2617154027247e43d32faff6a5  csv2xlsx_amd64.exe
    49b1ed81c1d3dc15ef24618a97892bda91216103466db3cfb8b8811c7cf5ed33  csv2xlsx_linux_386
    d4a01f8ae47c7c6315e828df06070e93202b4448e98bf957799d3be389be7209  csv2xlsx_linux_amd64
    dcbdb99e29552afcd26755a0bbd993180e8e1ad5ce21bd6f8f351cebed9bf6c5  csv2xlsx_osx


### Usage

You execute the program from a command line shell and prove at least the input file and the output file name.
Please see below for a list of command line options.

### Command line options

  -?	display usage information
  -colsep string
    	column separator (default '|')  (default "|")
  -columns string
    	column range to use (see below)
  -dateformat string
    	format for date cells (default YYYY-MM-DD) (default "2006-01-02")
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
        Each comma group can take a type specifiers for the column, one of "text", "number", "date" or "standard",
        separated from numbers with a colon (e.g. 0:text,3-16:number,17:date)


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
