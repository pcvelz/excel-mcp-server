package main

import (
	"fmt"
	"os"

	"github.com/negokaz/excel-mcp-server/internal/server"
)

func main() {
	s := server.New()
	err := s.Start()
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to start the server: %v\n", err)
		os.Exit(1)
	}
}
