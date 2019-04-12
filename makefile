
VERSION=`git describe --tags`
BUILD=`date +%FT%T%z`

# Setup the -ldflags option for go build here, interpolate the variable values
LDFLAGS=-ldflags "-X main.versionInfo=${VERSION} -X main.buildTimestamp=${BUILD}"

# Builds the project
build: ; go build ${LDFLAGS}

# Installs our project: copies binaries
install: ; go install ${LDFLAGS}

# cross compile
platforms:
	GOOS=windows GOARCH=386 go build ${LDFLAGS} -o csv2xlsx_386.exe
	GOOS=windows GOARCH=amd64 go build ${LDFLAGS} -o csv2xlsx_amd64.exe
	GOOS=linux GOARCH=386 go build ${LDFLAGS} -o csv2xlsx_linux_386
	GOOS=linux GOARCH=amd64 go build ${LDFLAGS} -o csv2xlsx_linux_amd64
	go build ${LDFLAGS} -o csv2xlsx_osx


# Cleans our project: deletes binaries
clean: ; if [ -f ${BINARY} ] ; then rm ${BINARY} ; fi

