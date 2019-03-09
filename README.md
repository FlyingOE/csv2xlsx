## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
[LICENSE](./LICENSE) file), you may download the binary.

Here are the SHA-256 checksums for the binaries:

	c0e1fe8195de7a144ea08973f6f4000a38cb7eae26c922b321828adcac6ed19e  csv2xlsx_386.exe
	0200f788d431e25d5610031ec63421b9d2f5f0fd2da26860bb6505362f71ace9  csv2xlsx_amd64.exe
	2e2b219ad56ee28d8c43f3efaec1d95385cbcab9b3d67a0db64271045a6624aa  csv2xlsx_linux_386
	9afc7eb3909e76a27d7b6fb82cc4d1b420c4f4140291f3396277522e7781334e  csv2xlsx_linux_amd64
	8b15fd767949a3ca457e503670e6fc89dc263bf3e7af0e15201b8c2b529cbde2  csv2xlsx_osx


### Usage

You execute the program from the command prompt of your operating system.

Ths most basic use case is `csv2xlsx -infile test.csv -outfile result.xlsx`, where you
take an input CSV file in UTF-8 and write out the .xslx file under a new file name.

To list all available options, start `csv2xlsx` with the option `--help`.

To list all supported encodings, execute `csv2xlsx` with the option `--listencodings`

### Source

This tool fulfills a special requirement and I will extend its functionality, if need arises. As I found out there are lots 
of people looking for such a tool, I decided to make it publicly available. I am in the process of learning Go and therefore
I am sure there are much better, more Go-idiomatic ways to achieve this functionality. If you have feedback on how to improve
the code or want to contribute, please do not hesitate to do so. I'd really like to improve my GO skills and learn things.
As my spare time for coding is limited to some hours around midnight a week, so please have some patience with my answers.
I am still amazed what you can accomplish within such a small tool in terms of making the admin part of my life easier. :-)


