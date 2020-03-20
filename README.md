## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
[LICENSE](./LICENSE) file), you may download the binary.

Here are the SHA-256 checksums for the latest binaries:

    eac04939a4e9d31da15493bd96c8f53ae685b3232e684be1f725a4c21b41f311  csv2xlsx_386.exe
    c385881df48fb1105a30c6e7ce3673c37587f9e7268ef4baa2e513d9dcfed70f  csv2xlsx_amd64.exe
    fbf9f4d334cb48559b4e3c7486907024862cdc44b2e2a1a6a7389a810aac7283  csv2xlsx_linux_386
    d4acf3dca4c96e54f10e00a7b075cdba4a7617a29e1b28c63f697daced6c202b  csv2xlsx_linux_amd64
    56284827897d87917c76ea8adab3c2319eef2dbd4531cba38399d92f7b730731  csv2xlsx_osx

### Usage

You execute the program from the command prompt of your operating system.

Ths most basic use case is `csv2xlsx --infile test.csv --outfile result.xlsx`, where you
take an input CSV file in UTF-8 and write out the .xslx file under a new file name.

To list all available options, start `csv2xlsx` with the option `--help`.

To list all supported encodings, execute `csv2xlsx` with the option `--listencodings`

There is no difference if you use one or two hyphens before an option (`-infile` is the same as `--infile`)

### Source

This tool fulfills a special requirement and I will extend its functionality, if need arises. As I found out there are lots 
of people looking for such a tool, I decided to make it publicly available. I am in the process of learning Go and therefore
I am sure there are much better, more Go-idiomatic ways to achieve this functionality. If you have feedback on how to improve
the code or want to contribute, please do not hesitate to do so. I'd really like to improve my GO skills and learn things.
As my spare time for coding is limited to some hours around midnight a week, so please have some patience with my answers.
I am still amazed what you can accomplish within such a small tool in terms of making the admin part of my life easier. :-)


