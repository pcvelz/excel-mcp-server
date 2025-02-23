package mcp

import (
	"fmt"

	"github.com/mark3labs/mcp-go/mcp"
)

func NewToolResultInvalidArgumentError(message string) *mcp.CallToolResult {
	return mcp.NewToolResultError(fmt.Sprintf("Invalid argument: %s", message))
}
