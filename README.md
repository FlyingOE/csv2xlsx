## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
[LICENSE](./LICENSE) file), you may download the binary.

Here are the SHA-256 checksums for the latest binaries:

	1283c1be065ef5e3f100dffc1149389de4d5cd61d976c4f8f44bb5c69d03b9d0  csv2xlsx_386.exe
	fc0b50435cca22ed2413c0e000eeb76072cba90a712b7b27a2c5896946060e4e  csv2xlsx_amd64.exe
	9bb19bb0b85f4157933e71671a6761f9ca9261cac700812bcece75b97549c6bf  csv2xlsx_linux_386
	d889df79c48334a5930ace16a2e9d3873eede2b4854cbd7470134cd373f2d3fc  csv2xlsx_linux_amd64
	07152f406454a13192d75ca38a1fd36c7818a7501c46ae35d8f3b4c58f82cdb0  csv2xlsx_osx

### Usage

You execute the program from the command prompt of your operating system.

Ths most basic use case is `csv2xlsx --infile test.csv --outfile result.xlsx`, where you
take an input CSV file in UTF-8 and write out the .xslx file under a new file name.

To list all available options, start `csv2xlsx` with the option `--help`.

To list all supported encodings, execute `csv2xlsx` with the option `--listencodings`

There is no difference if you use one or two hyphens before an option (`-infile` is the same as `--infile`)

#### Default column and row separators

Please note that the **default column separator** is the pipe char (`|`) and the **default row separator** is the newline char (`\n`). 
The tools came into existence to solve a problem for me, so this is the default you will have to live with or use the `--colsep` and `--rowsep` parameters. ;-)


### Source

This tool fulfills a special requirement and I will extend its functionality, if need arises. As I found out there are lots 
of people looking for such a tool, I decided to make it publicly available. I am in the process of learning Go and therefore
I am sure there are much better, more Go-idiomatic ways to achieve this functionality. If you have feedback on how to improve
the code or want to contribute, please do not hesitate to do so. I'd really like to improve my GO skills and learn things.
As my spare time for coding is limited to some hours around midnight a week, so please have some patience with my answers.
I am still amazed what you can accomplish within such a small tool in terms of making the admin part of my life easier. :-)


