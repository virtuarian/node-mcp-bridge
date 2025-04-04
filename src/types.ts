import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { FSWatcher } from "chokidar";

// MCP Server Configuration
export interface ServerConfig {
  command: string;
  args?: string[];
  env?: Record<string, string>;
  autoApprove?: string[];
  disabled?: boolean;
  timeout?: number;
  sessionTimeout?: number; // undefined: default
}

// MCP Server Information
export interface McpServer {
  name: string;
  config: string; // JSON string
  status: 'connected' | 'connecting' | 'disconnected';
  error: string;
  disabled: boolean;
  tools: McpTool[];
  resources: McpResource[];
  resourceTemplates: McpResourceTemplate[];
}

// MCP Connection Information
export interface McpConnection {
  server: McpServer;
  client: Client | null;
  transport: StdioClientTransport | null;
}


// MCP Tool Information
export interface McpTool {
  serverName: string;
  name: string;
  description: string;
  schema?: Record<string, any>;
  autoApprove?: boolean;
}

// MCP Resource Information
export interface McpResource {
  uri: string;
  description: string;
}

// MCP Resource Template Information
export interface McpResourceTemplate {
  kind: string;
  description: string;
}

// MCP Tool Call Response
export interface McpToolCallResponse {
  result: any;
}

// MCP Resource Read Response
export interface McpResourceResponse {
  content: string;
  mediaType: string;
}

// MCP Session Settings
export interface McpSettings {
  mcpServers: Record<string, ServerConfig>;
}

// Server Manager State
export interface ServerManagerState {
  connections: Map<string, McpConnection>;
  fileWatchers: Map<string, FSWatcher>;
  isConnecting: boolean;
}

// MCP Session State
export interface SessionState {
  id: string;
  approvedTools: Map<string, Set<string>>; // ServerName -> TooName
  createdAt: Date;
  lastActive: Date;
}

// MCP Session Manager State
export interface ApprovalRequiredResponse {
  error: string;
  approvalRequired: true;
  serverName: string;
  toolName: string;
  toolDescription: string;
  sessionId: string;
}