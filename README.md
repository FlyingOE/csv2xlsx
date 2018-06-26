## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
license below), you may download the binary.

Here are the SHA-256 checksums for the binaries:

    86fab8cdb756d612391bdfca36641414424cb0cfe7c9c196329124976f3d3a8c  csv2xlsx_386.exe
    91b94bb4c0acf91bcd2b3874d7ab7f96e204e6c0acb1d0119694bb740dedb6f4  csv2xlsx_amd64.exe
    d1b3dc8bfa72647f4e92dbafede0ee729ccb488a7b2a400304634bf03439b744  csv2xlsx_linux_386
    3e8661e7ef681c796452736e9004f19653ccf01c916f3c6a8b1e67d99f1e0ab5  csv2xlsx_linux_amd64
    2933cdca783beca8fbcfccc2d396f4ec115c898a9f69680d6c64806ac84e1804  csv2xlsx_osx    

### Usage

You execute the program from the command prompt of your operating system.

Ths most basic use case is `csv2xlsx -infile test.csv -outfile result.xlsx`, where you
take an input CSV file in UTF-8 and write out the .xslx file under a new file name.

To list all available options start `csv2xlsx` with the option `--help`.

### Supported encodings

    Codepage037      Codepage437      Codepage850      Codepage852      Codepage855      
    Codepage858      Codepage860      Codepage862      Codepage863      Codepage865
    Codepage866      Codepage1047     Codepage1140      
    ISO8859_1        ISO8859_2        ISO8859_3        ISO8859_4        ISO8859_5
    ISO8859_6        ISO8859_6E       ISO8859_6I       ISO8859_7        ISO8859_8
    ISO8859_8E       ISO8859_8I       ISO8859_9        ISO8859_10       ISO8859_13
    ISO8859_14       ISO8859_15       ISO8859_16      
    Koi8r            Koi8u      
    Macintosh        MacintoshCyrillic      
    Windows874       Windows1250      Windows1251      Windows1252      Windows1253      
    Windows1254      Windows1255      Windows1256      Windows1257      Windows1258      


### Source

This tool fulfills a special requirement and I will extend its functionality, if need arises. As I found out there are lots 
of people looking for such a tool, I decided to make it publicly available. I am in the process of learning Go and therefore
I am sure there are much better, more Go-idiomatic ways to achieve this functionality. If you have feedback on how to improve
the code or want to contribute, please do not hesitate to do so. I'd really like to improve my GO skills and learn things.
As my spare time for coding is limited to some hours around midnight a week, so please have some patience with my answers.
I am still amazed what you can accomplish within less than 500 lines of code in terms of making my admin part of life easier. :-)


