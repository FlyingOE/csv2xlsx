## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
license below), you may download the binary.

Here are the SHA-256 checksums for the binaries:

    63793d7f4ed54050611237c69f0030bd6df3d20804f4307cea8a532d5b63fcda  csv2xlsx_386.exe
    a10ae55fe039f781f0b002da5db48b81928947ce24dcff73e95a43ac99c458d0  csv2xlsx_amd64.exe
    2e6d5181c3b7ffaefc6d14e8dd6dd273af7b9b41bad91bbde23c40aba2334c06  csv2xlsx_linux_386
    415e66ca0fcc602fb4dbd9eda070826b5c1cc38979685ddd72373be8f29001ff  csv2xlsx_linux_amd64
    812142075e9bf187a8a6581911ba565cfde600e59c26f3dcae2e2d021ff05550  csv2xlsx_osx


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


