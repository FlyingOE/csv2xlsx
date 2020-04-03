## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
[LICENSE](./LICENSE) file), you may download the binary.

Here are the SHA-256 checksums for the latest binaries:

    e6be9f41cf62281303c1480f3e0219eecaa01cfc79718a09d9d7ab79af3d5bca  csv2xlsx_386.exe
    866f6af6ef7f5f54d8b19758b0a551c10b9cedf3111624f2d5e1e964aff675f1  csv2xlsx_amd64.exe
    4a82669c599dd11e679822d9469c8fe7d326f2d23460cbda7331fec4ac4af0c2  csv2xlsx_linux_386
    b4cb8e8d41bde67ba102f796a40bfef128ab822a4c192d3dffe908db78925d91  csv2xlsx_linux_amd64
    82b0b1310b69a1ce03357caa47b02b80428a814f71896d64ce254b2e304590fd  csv2xlsx_osx

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


