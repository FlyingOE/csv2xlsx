## CSV2XLSX

Finally: a simple, single file executable, no runtime libs command line tool to convert
a CSV file to XLSX. And you may even select line and column ranges. :-)

This programm has been created due to an internal requirement for a Bash shell script. After searching
the web I found out that even in 2017 there is no simple, binary executable that does not need any
runtime, virtual machine or whatever. All you need is a compiler for the Go programming language.
If you do not want to compile the source and you decide to trust me (no warranty whatsoever, see the
[LICENSE](./LICENSE) file), you may download the binary.

Here are the SHA-256 checksums for the latest binaries:

eabdc8bfc9c23ad1e04e4a6576da5c80456f16ac3c870283a5571e09f8831826  csv2xlsx_386.exe
69dc4821dd4db4465a84f74fe592f145ba7574466290f60f803c24e1724161ed  csv2xlsx_amd64.exe
0de1192da0a1c217e018586a07aae4991e1c481d76e47e7d77b84958c21019bd  csv2xlsx_linux_386
0c2fa046995b2757998a8057c06f91c9c201e427246054faf2a4612d137ac889  csv2xlsx_linux_amd64
3e4178237567295d9c8f14577ad6733b722e87f4f5b7f90cf7bb73124a2f8fb1  csv2xlsx_osx

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


