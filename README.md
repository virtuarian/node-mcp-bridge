# Node MCP Bridge - README

## Overview

Node MCP Bridge is middleware that coordinates between Model Context Protocol (MCP) servers and clients. It centrally manages multiple MCP servers and routes client requests to the appropriate server. Key features include:

- Management of multiple MCP servers (Playwright, Puppeteer, FileSystem, etc.)
- RESTful API interface
- Session management and tool authorization flow
- Web-based administration interface
- Session timeout settings per server (default 180 minutes, unlimited option available)

## Installation

### Prerequisites

- Node.js 18 or higher

### Installation Steps

```bash
# Clone the repository
git clone https://github.com/virtuarian/node-mcp-bridge.git
cd node-mcp-bridge

# Install dependencies
npm install

# Build TypeScript
npm run build
```

## Quick Start

### Launch Procedure

```bash
# Development mode (watches for source code changes)
npm run dev

# Production mode
npm start
```

By default, the server starts at `http://localhost:3001`. The port can be changed in the .env file:

```
PORT=8080
```

## Calling from Other Applications

Node MCP Bridge provides a RESTful API. It can be called from other applications as follows:

```javascript
// Session ID (arbitrary ID)
const sessionId = 'test-sessionid';

// Tool call
const toolResponse = await fetch(`http://localhost:3001/tools/call/${sessionId}`, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    serverName: 'playwright',
    toolName: 'browser_navigate',
    arguments: { url: 'https://example.com' }
  })
});

const result = await toolResponse.json();
```

## API Specification

### Tool-related

| Endpoint | Method | Description |
|--------------|--------|------|
| `/tools` | GET | Get a list of all available tools |
| `/tools/call` | POST | Call tools without a session (auto-approved tools only) |
| `/tools/call/:sessionId` | POST | Call tools with a specific session |
| `/tools/call/:sessionId/approve` | POST | Approve and call a specific tool for a session |

### Admin API

| Endpoint | Method | Description |
|--------------|--------|------|
| `/admin/servers` | GET | Get a list of all registered servers |
| `/admin/servers/:serverName` | PUT | Add/update server configuration |
| `/admin/servers/:serverName` | DELETE | Delete a server |
| `/admin/servers/:serverName/restart` | POST | Restart a server |
| `/admin/servers/:serverName/toggleDisabled` | PUT | Toggle server enabled/disabled state |
| `/admin/servers/:serverName/tools` | GET | Get a list of server tools |
| `/admin/servers/:serverName/tools/:toolName/toggleAutoApprove` | PUT | Toggle auto-approval setting for a tool |

## Server Management

The server management interface is accessible at `http://localhost:3001/admin`.

### Using the Admin Interface

The admin interface allows the following operations:

1. **Adding a Server**: Click the "Add Server" button and enter the required information.
   - Server Name: A unique identifier
   - Command: The command to execute (e.g., `npx`)
   - Arguments: Specified one per line (e.g., `@playwright/mcp@latest`)
   - Environment Variables: Add as needed
   - Timeout: Operation timeout in seconds
   - Session Timeout: Session validity period in minutes, 0 for unlimited

2. **Server Management**:
   - Details: View detailed server information
   - Restart: Restart the server
   - Edit: Edit server configuration
   - Delete: Delete the server
   - Enable/Disable: Toggle switch to change server state

3. **Tool Management**:
   - Toggle auto-approval settings for each tool on the server details page

### Language Settings

Select the admin interface language from the dropdown menu in the top right:
- English
- 日本語 (Japanese)
- 中文 (Chinese)

## Examples

### Calling from Excel

Example of calling Node MCP Bridge from Excel macros
- Sample Excel file is available at call_excelmacro.xlsm.

```vb
Sub CallMcpBridge()
    Dim sessionId As String
    Dim response As String
    Dim payload As String
    Dim serverName As String
    Dim toolName As String
    Dim endPoint As String
    
    ' Session creation
    ' Set an arbitrary session ID
    sessionId = "abcd"
    
    ' Tool call
    ' EndPoint
    endPoint =  "http://localhost:3001/tools/call/" & sessionId & "/approve"

    ' Open https://example.com with playwright
    serverName = "playwright"
    toolName = "browser_navigate"

    payload = "{""serverName"":""" & serverName & """,""toolName"":""" & toolName & """,""arguments"":{""url"":""https://example.com""}}"

    ' Send request    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "POST", endPoint, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.send payload

    ' Get response
    response = httpRequest.responseText
    
    MsgBox "Response: " & response
End Sub
```

### Calling from Next.js

Example of calling Node MCP Bridge from a Next.js application

```javascript
// pages/api/mcp-bridge.js
export default async function handler(req, res) {
  try {
    // Set an arbitrary session ID (use an appropriate identifier in actual apps)
    const sessionId = 'user-session-' + Math.random().toString(36).substring(2, 10);
    
    // Tool call
    // Save as PDF after navigating to a page in a previous process
    const toolRes = await fetch(`http://localhost:3001/tools/call/${sessionId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        serverName: 'playwright',
        toolName: 'browser_save_as_pdf',
        arguments: {}
      })
    });
    
    const response = await toolRes.json();
    
    // Handle approval if needed
    if (response.approvalRequired) {
      console.log('Tool requires approval, sending approval request');
      
      // Call approval endpoint
      const approvalRes = await fetch(`http://localhost:3001/tools/call/${sessionId}/approve`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          serverName: response.serverName,
          toolName: response.toolName,
          arguments: {}
        })
      });
      
      const result = await approvalRes.json();
      res.status(200).json(result);
    } 
    else {
      // Auto-approved or already approved
      res.status(200).json(response);
    }
  } 
  catch (error) {
    res.status(500).json({ error: error.message });
  }
}
```

Frontend usage example:

```javascript
// pages/index.js
import { useState } from 'react';

export default function Home() {
  const [pdfInfo, setPdfInfo] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  
  const savePdf = async () => {
    setLoading(true);
    setError(null);
    
    try {
      const res = await fetch('/api/mcp-bridge');
      const data = await res.json();
      
      if (data.error) {
        setError(data.error);
      } 
      else if (data.result && data.result.content) {
        // Display PDF save result (typically returns a text message)
        // Example: "Saved as C:\\Users\\user\\AppData\\Local\\Temp\\page-2025-04-04T07-43-22-385Z.pdf"
        const textResult = data.result.content.find(item => item.type === 'text')?.text || '';
        setPdfInfo(textResult);
      }
    } 
    catch (error) {
      console.error('Error:', error);
      setError(error.message);
    } 
    finally {
      setLoading(false);
    }
  };
  
  return (
    <div className="container mx-auto p-4">
      <h1 className="text-2xl font-bold mb-4">Save PDF</h1>
      
      <button 
        className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded"
        onClick={savePdf} 
        disabled={loading}
      >
        {loading ? 'Processing...' : 'Save PDF'}
      </button>
      
      {error && (
        <div className="mt-4 p-3 bg-red-100 text-red-700 rounded">
          Error: {error}
        </div>
      )}
      
      {pdfInfo && (
        <div className="mt-4">
          <h2 className="text-xl font-semibold mb-2">Result:</h2>
          <div className="p-4 bg-gray-100 rounded">
            <p>{pdfInfo}</p>
          </div>
        </div>
      )}
    </div>
  );
}
```