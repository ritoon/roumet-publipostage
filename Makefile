# bin/bash

build-w:
	GOOS=windows GOARCH=amd64 CGO_ENABLED=0 go build -o generate-doc.exe .