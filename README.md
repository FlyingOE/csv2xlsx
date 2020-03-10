## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
[LICENSE](./LICENSE) file), you may download the binary.

Here are the SHA-256 checksums for the binaries:

	bd05f824315c104964307cd6b8526f8b6d2b256f53faa55f4fee24fb39896c16  csv2xlsx_386.exe
	7eec305c198bd3a10bc0647ae9614d51e1b01e906100a65ba3fbbdd5566e8ff3  csv2xlsx_amd64.exe
	d46e16487cdc4334055390921b1994387e49c33e8e9556828b5a942373fcc480  csv2xlsx_linux_386
	361a78d2650556c14669460ab0c656c0a7dba949ebf2846be37d2ddd719c0964  csv2xlsx_linux_amd64
	186706ac7860f16be50363b41043489a8b838bd909bb9981ff5e0fb84baac091  csv2xlsx_osx


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


