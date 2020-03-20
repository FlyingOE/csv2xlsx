## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
[LICENSE](./LICENSE) file), you may download the binary.

Here are the SHA-256 checksums for the latest binaries:

    1be9fe928da80e9d3b0fdb95c9712ac410c70c8a8e85d5f416852ce9731f067e  csv2xlsx_386.exe
    46ff275e48d7b25778b0d7878aa1140507c3b3b0d3a1ba6bba9e7125cbb4d095  csv2xlsx_amd64.exe
    80e4181c4a129b78988fb707fae4f5b00ae140cf4457bf145e25d27e9db15f7f  csv2xlsx_linux_386
    26393bc44e51356534c42406c770647fc4c3bf0c07c19ae4f3e4d04de1b42d5b  csv2xlsx_linux_amd64
    fd6f3943647a6f355137ac1f916eaa1dd96e00689e4db7d66b9eb7a12aa2071c  csv2xlsx_osx

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


