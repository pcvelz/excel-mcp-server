#!/usr/bin/env node
"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
const path = __importStar(require("path"));
const childProcess = __importStar(require("child_process"));
const BINARY_DISTRIBUTION_PACKAGES = {
    win32_ia32: "excel-mcp-server_windows_386_sse2",
    win32_x64: "excel-mcp-server_windows_amd64_v1",
    win32_arm64: "excel-mcp-server_windows_arm64_v8.0",
    darwin_x64: "excel-mcp-server_darwin_amd64_v1",
    darwin_arm64: "excel-mcp-server_darwin_arm64_v8.0",
    linux_ia32: "excel-mcp-server_linux_386_sse2",
    linux_x64: "excel-mcp-server_linux_amd64_v1",
    linux_arm64: "excel-mcp-server_linux_arm64_v8.0",
};
function getBinaryPath() {
    const suffix = process.platform === 'win32' ? '.exe' : '';
    const pkg = BINARY_DISTRIBUTION_PACKAGES[`${process.platform}_${process.arch}`];
    if (pkg) {
        return path.resolve(__dirname, pkg, `excel-mcp-server${suffix}`);
    }
    else {
        throw new Error(`Unsupported platform: ${process.platform}_${process.arch}`);
    }
}
childProcess.execFileSync(getBinaryPath(), process.argv, {
    stdio: 'inherit',
});
