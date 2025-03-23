package server

import (
	"runtime"

	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/tools"
)

type ExcelServer struct {
	server *server.MCPServer
}

func New(version string) *ExcelServer {
	s := &ExcelServer{}
	s.server = server.NewMCPServer(
		"excel-mcp-server",
		version,
	)
	tools.AddReadSheetNamesTool(s.server)
	tools.AddReadSheetDataTool(s.server)
	tools.AddReadSheetFormulaTool(s.server)
	if runtime.GOOS == "windows" {
		tools.AddReadSheetImageTool(s.server)
	}
	tools.AddWriteSheetDataTool(s.server)
	tools.AddWriteSheetFormulaTool(s.server)
	return s
}

func (s *ExcelServer) Start() error {
	return server.ServeStdio(s.server)
}
