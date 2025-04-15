/**
 * Copyright 2025 virtuarian
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import chokidar from 'chokidar';
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import {
  CallToolResultSchema,
  ListResourcesResultSchema,
  ListResourceTemplatesResultSchema,
  ListToolsResultSchema,
  ReadResourceResultSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { setTimeout as setTimeoutPromise } from 'timers/promises';
import logger from './logger.js';
import { config, loadMcpSettings, saveMcpSettings, secondsToMs, validateServerConfig } from './config.js';
import {
  McpConnection,
  ServerConfig,
  McpTool,
  McpResource,
  McpResourceTemplate,
  McpToolCallResponse,
  McpResourceResponse,
  McpServer,
  ServerManagerState
} from './types.js';

export class MCPServerManager {
  private connections: Map<string, McpConnection> = new Map();
  private fileWatchers: Map<string, chokidar.FSWatcher> = new Map();
  private isConnecting: boolean = false;

  constructor() {
    logger.info('Initializing MCP Server Manager');
  }

  // Create MCP connection structure
  private createConnectionObject(name: string, serverConfig: ServerConfig): McpConnection {
    return {
      server: {
        name,
        config: JSON.stringify(serverConfig),
        status: 'disconnected',
        error: '',
        disabled: serverConfig.disabled || false,
        tools: [],
        resources: [],
        resourceTemplates: [],
      },
      client: null,
      transport: null,
    };
  }

  // Connect to MCP server
  async connectToServer(name: string, serverConfig: ServerConfig): Promise<McpConnection> {
    const log = logger.child({ serverName: name });

    // Delete existing connection if any
    if (this.connections.has(name)) {
      await this.deleteConnection(name);
    }

    try {
      log.info('Connecting to MCP server');

      // Create MCP client
      const client = new Client(
        {
          name: "NodeMCPBridge",
          version: "1.0.0",
        },
        {
          capabilities: {},
        }
      );

      // Configure transport
      const transport = new StdioClientTransport({
        command: serverConfig.command,
        args: serverConfig.args || [],
        env: {
          ...serverConfig.env,
          ...(process.env.PATH ? { PATH: process.env.PATH } : {}),
        },
        stderr: "pipe", // Capture stderr
      });

      // Create connection object
      const connection = this.createConnectionObject(name, serverConfig);
      connection.client = client;
      connection.transport = transport;
      connection.server.status = 'connecting';

      this.connections.set(name, connection);

      // Setup error handler
      transport.onerror = async (error: Error) => {
        log.error({ error: error.message }, 'Transport error occurred');
        const conn = this.connections.get(name);
        if (conn) {
          conn.server.status = 'disconnected';
          this.appendErrorMessage(conn, error.message);
        }
      };

      // Setup close handler
      transport.onclose = async () => {
        log.info('Transport connection closed');
        const conn = this.connections.get(name);
        if (conn) {
          conn.server.status = 'disconnected';
        }
      };

      // Validate configuration
      if (!validateServerConfig(serverConfig)) {
        throw new Error('Invalid Server Configuration');
      }

      // Start transport first to capture stderr
      await transport.start();
      const stderrStream = transport.stderr;
      if (stderrStream && typeof stderrStream.on === 'function') {
        let errorBuffer = '';
        const maxErrorLength = 5000; // 最大エラーメッセージ長を制限

        stderrStream.on('data', (data: Buffer) => {
          try {
            const errorOutput = data.toString().trim();
            if (!errorOutput) return;

            // Log the error output
            log.debug({ stderr: errorOutput }, 'Server stderr output');

            // Append to error buffer
            errorBuffer += errorOutput + '\n';

            // Truncate if too long
            if (errorBuffer.length > maxErrorLength) {
              const truncateMsg = '\n[... Some messages truncated ...]\n';
              errorBuffer = errorBuffer.substring(errorBuffer.length - maxErrorLength + truncateMsg.length);
              errorBuffer = truncateMsg + errorBuffer;
            }

            // Update connection error message
            const conn = this.connections.get(name);
            if (conn) {
              conn.server.error = errorBuffer;
            }
          } catch (e) {
            log.error({ error: e instanceof Error ? e.message : String(e) }, 'Error processing stderr data');
          }
        });
      } else {
        log.warn('Stderr stream not available or invalid');
      }

      // Replace start() method with no-op (to prevent double start in connect())
      transport.start = async () => { };

      // Establish connection
      await client.connect(transport);
      connection.server.status = 'connected';
      connection.server.error = '';

      // Fetch initial data after connection
      connection.server.tools = await this.fetchToolsList(name);
      connection.server.resources = await this.fetchResourcesList(name);
      connection.server.resourceTemplates = await this.fetchResourceTemplatesList(name);

      log.info({
        toolsCount: connection.server.tools.length,
        resourcesCount: connection.server.resources.length,
        templatesCount: connection.server.resourceTemplates.length
      }, 'Connected to MCP server successfully');

      // Setup file watcher
      this.setupFileWatcher(name, serverConfig);

      return connection;
    } catch (error) {
      // Error handling
      const connection = this.connections.get(name);
      if (connection) {
        connection.server.status = 'disconnected';
        this.appendErrorMessage(connection, error instanceof Error ? error.message : String(error));
      }
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Connection failed');
      throw error;
    }
  }

  // Append error message
  private appendErrorMessage(connection: McpConnection, error: string): void {
    const newError = connection.server.error ? `${connection.server.error}\n${error}` : error;
    connection.server.error = newError;
  }

  // Disconnect from MCP server
  async deleteConnection(name: string): Promise<void> {
    const log = logger.child({ serverName: name });
    const connection = this.connections.get(name);
    if (!connection) return;

    try {
      log.info('Deleting server connection');

      // Delete file watcher
      const watcher = this.fileWatchers.get(name);
      if (watcher) {
        watcher.close();
        this.fileWatchers.delete(name);
      }

      // Close connection
      if (connection.transport) {
        await connection.transport.close();
      }
      if (connection.client) {
        await connection.client.close();
      }

      this.connections.delete(name);
      log.info('Connection deleted successfully');
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Disconnect error');
    }
  }

  // Setup file watcher
  private setupFileWatcher(name: string, serverConfig: ServerConfig): void {
    const log = logger.child({ serverName: name });

    // Monitor file if path is found
    const filePath = serverConfig.args?.find(arg => arg.includes('build/index.js') || arg.endsWith('.js'));
    if (filePath) {
      log.debug({ filePath }, 'Setting up file watcher');

      // Close existing watcher if any
      const existingWatcher = this.fileWatchers.get(name);
      if (existingWatcher) {
        existingWatcher.close();
      }

      // Setup new watcher
      const watcher = chokidar.watch(filePath, {
        persistent: true,
        ignoreInitial: true,
        awaitWriteFinish: true,
      });

      watcher.on('change', () => {
        log.info({ filePath }, 'File change detected, restarting server');
        this.restartConnection(name);
      });

      this.fileWatchers.set(name, watcher);
      log.info({ filePath }, 'File monitoring started');
    }
  }

  // Restart MCP server connection
  async restartConnection(name: string): Promise<void> {
    const log = logger.child({ serverName: name });
    this.isConnecting = true;

    const connection = this.connections.get(name);
    if (!connection) {
      log.warn('Connection not found, cannot restart');
      this.isConnecting = false;
      return;
    }

    const serverConfig = connection.server.config;
    if (!serverConfig) {
      log.warn('Server configuration not found, cannot restart');
      this.isConnecting = false;
      return;
    }

    log.info('Restarting MCP server');
    connection.server.status = 'connecting';
    connection.server.error = '';

    try {
      // Add a short delay for humans to recognize changes
      await setTimeoutPromise(500);

      // Disconnect and reconnect
      await this.deleteConnection(name);
      await this.connectToServer(name, JSON.parse(serverConfig));

      log.info('MCP server restarted successfully');
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Restart error');
    }

    this.isConnecting = false;
  }

  // Update server connections
  async updateServerConnections(newServers: Record<string, ServerConfig>): Promise<void> {
    logger.info({ serverCount: Object.keys(newServers).length }, 'Updating server connections');
    this.isConnecting = true;

    try {
      // Create sets of current and new server names
      const currentNames = new Set(this.connections.keys());
      const newNames = new Set(Object.keys(newServers));

      // Disconnect removed servers
      for (const name of currentNames) {
        if (!newNames.has(name)) {
          logger.info({ serverName: name }, 'Removing server that no longer exists');
          await this.deleteConnection(name);
        }
      }

      // Connect new or changed servers
      for (const [name, serverConfig] of Object.entries(newServers)) {
        const log = logger.child({ serverName: name });
        const currentConnection = this.connections.get(name);

        if (!currentConnection) {
          // New server
          log.info('Adding new server');
          try {
            await this.connectToServer(name, serverConfig);
          } catch (error) {
            log.error({ error: error instanceof Error ? error.message : String(error) }, 'New server connection failed');
          }
        } else {
          // Existing server - check if configuration changed
          const currentConfig = JSON.parse(currentConnection.server.config);
          const configChanged = JSON.stringify(currentConfig) !== JSON.stringify(serverConfig);

          if (configChanged) {
            log.info('Configuration changed, reconnecting server');
            try {
              await this.deleteConnection(name);
              await this.connectToServer(name, serverConfig);
            } catch (error) {
              log.error({ error: error instanceof Error ? error.message : String(error) }, 'Reconnection failed');
            }
          }
        }
      }
    } catch (error) {
      logger.error({ error: error instanceof Error ? error.message : String(error) }, 'Server connection update failed');
    }

    this.isConnecting = false;
    logger.info('Server connections update completed');
  }

  // Fetch tools list from all servers
  async fetchAllToolsList(): Promise<McpTool[]> {
    logger.info('Fetching tools from all servers');
    const servers = this.getAllServers().filter(server => !server.disabled);
    const allTools: McpTool[] = [];

    for (const server of servers) {
      const tools = await this.fetchToolsList(server.name);
      allTools.push(...tools);
    }

    logger.info({ toolCount: allTools.length }, 'All tools fetched successfully');
    return allTools;
  }

  // Fetch tools list
  async fetchToolsList(serverName: string): Promise<McpTool[]> {
    const log = logger.child({ serverName });
    log.debug('Fetching tools list');

    try {
      const connection = this.connections.get(serverName);
      if (!connection || !connection.client) {
        log.warn('Connection not available, returning empty tools list');
        return [];
      }

      // Check if server is disabled
      if (connection.server.disabled) {
        log.info('Server is disabled, returning empty tools list');
        return [];
      }

      // Fetch tools from server
      const response = await connection.client.request(
        { method: "tools/list" },
        ListToolsResultSchema
      );

      // Apply autoApprove settings
      const serverConfig = JSON.parse(connection.server.config) as ServerConfig;
      const autoApproveList = serverConfig.autoApprove || [];

      // Add autoApprove flag to tools
      const tools = (response?.tools || []).map(tool => {
        log.debug({ toolName: tool.name }, 'Tool information retrieved');

        // Normalize schema information and return tool info
        interface JSONSchema {
          type: string;
          properties: Record<string, unknown>;
          required: string[];
          additionalProperties?: boolean;
        }

        const schema = (tool.inputSchema || {
          type: 'object',
          properties: {},
          required: []
        }) as unknown as JSONSchema;

        // Ensure schema consistency
        if (schema.type !== 'object') {
          schema.type = 'object';
        }

        if (!schema.properties) {
          schema.properties = {};
        }

        if (!schema.required) {
          schema.required = [];
        }

        return {
          serverName: serverName,
          name: tool.name,
          description: tool.description || '', // Ensure description is always a string
          schema: tool.inputSchema,
          autoApprove: autoApproveList.includes(tool.name)
        };
      });

      log.info({ toolCount: tools.length }, 'Tools fetched successfully');
      return tools;
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Failed to fetch tools list');
      return [];
    }
  }


  /**
   * Fetch tools for function calling format
   * @param serverName Optional server name to fetch tools from
   * @returns Array of tools in function calling format
   */
  async fetchFunctionCallingTools(serverName?: string): Promise<Array<{
    type: string;
    function: {
      name: string;
      description: string;
      parameters: Record<string, any>;
    };
  }>> {
    const log = logger.child({ serverName });
    log.info('Fetching tools for function calling format');

    try {

      // serverName is specified, check if the server is disabled
      if (serverName) {
        const connection = this.connections.get(serverName);
        if (connection?.server.disabled) {
          log.info('Specified server is disabled, returning empty tools list');
          return [];
        }
      }

      // Fetch tools list
      const mcpTools = serverName
        ? await this.fetchToolsList(serverName)
        : await this.fetchAllToolsList();

      // Filter out tools that are not auto-approvable
      const functionTools = mcpTools.map(tool => {
        const schema = tool.schema || {
          type: 'object',
          properties: {},
          required: []
        };

        const functionName = `${tool.serverName}__${tool.name}`
          .replace(/[^a-zA-Z0-9_.-]/g, '_')
          .toLowerCase();

        const description = tool.description ||
          `Tool '${tool.name}' from MCP server '${tool.serverName}'`;

        return {
          type: "function",
          function: {
            name: functionName,
            description: description,
            parameters: {
              ...schema,
              // 追加メタデータを含める（オプション）
              "x-mcp-metadata": {
                serverName: tool.serverName,
                toolName: tool.name,
                autoApprove: tool.autoApprove
              }
            }
          }
        };
      });

      log.info({ toolCount: functionTools.length }, 'Function calling tools prepared successfully');
      return functionTools;
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      log.error({ error: errorMsg }, 'Failed to fetch function calling tools');
      return [];
    }
  }

  /**
   * OpenAI Function Callingの呼び出し結果を実際のMCPツール呼び出しに変換する
   * @param functionCall OpenAIから返された関数呼び出し情報
   * @returns MCP Tool呼び出し結果
   */
  async executeFunctionCall(functionCall: {
    name: string;
    arguments: string;
  }): Promise<McpToolCallResponse> {
    const log = logger.child({ functionCall: functionCall.name });
    log.info('Executing function call from OpenAI');

    try {
      // 関数名からサーバー名とツール名を抽出
      const nameParts = functionCall.name.split('__');
      const serverName = nameParts[0];
      const toolName = nameParts.length > 1 ? nameParts[1] : functionCall.name;
  
      // 引数を解析
      let toolArguments: Record<string, any> = {};
      try {
        toolArguments = JSON.parse(functionCall.arguments);
      } catch (e) {
        log.error({ error: e instanceof Error ? e.message : String(e) }, 'Failed to parse function arguments');
        throw new Error('Invalid function arguments format');
      }

      // MCPツールを呼び出す
      log.info({ serverName, toolName }, 'Calling MCP tool from function call');
      return await this.callTool(serverName, toolName, toolArguments);
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      log.error({ error: errorMsg }, 'Function call execution failed');
      throw error;
    }
  }


  /**
 * Gemini Function Calling用にフォーマットされたツールリストを取得する
 * Gemini AIモデルのツール呼び出し機能と互換性のある形式でMCPツールを返す
 * 有効なサーバーからのみツールを取得
 * @param serverName ツールを取得するサーバー名（指定なしの場合は全有効サーバー）
 * @returns Gemini Function Calling形式のツール定義配列
 */
  async fetchGeminiFunctionCallingTools(serverName?: string): Promise<Array<{
    function_declarations: [{
      name: string;
      description: string;
      parameters: {
        type: "OBJECT";
        properties: Record<string, any>;
        required?: string[];
      };
    }]
  }>> {
    const log = logger.child({ serverName });
    log.info('Fetching tools for Gemini function calling format from enabled servers');

    try {
      // 特定のサーバーが指定された場合、そのサーバーが有効かチェック
      if (serverName) {
        const connection = this.connections.get(serverName);
        if (connection?.server.disabled) {
          log.info('Specified server is disabled, returning empty tools list');
          return [];
        }
      }

      // 全サーバーまたは特定のサーバーからツールを取得
      const mcpTools = serverName
        ? await this.fetchToolsList(serverName)
        : await this.fetchAllToolsList();

      // Gemini形式に変換
      const geminiTools = mcpTools.map(tool => {
        // MCPツールのスキーマ取得
        const schema = tool.schema || {
          type: 'object',
          properties: {},
          required: []
        };

        // 名前の作成（サーバー名とツール名を組み合わせる）- Gemini用
        const functionName = `${tool.serverName}__${tool.name}`
          .replace(/[^a-zA-Z0-9_.-]/g, '_')
          .toLowerCase();

        // 説明がない場合はデフォルトの説明を提供
        const description = tool.description ||
          `Tool '${tool.name}' from MCP server '${tool.serverName}'`;

        // GeminiのFunction Calling形式に変換
        // Geminiの場合はtypeが"OBJECT"（大文字）であることと、
        // プロパティ構造がOpenAIとは異なることに注意
        return {
          function_declarations: [{
            name: functionName,
            description: description,
            parameters: {
              type: "OBJECT" as const, // Using a const assertion to ensure type is the literal "OBJECT"
              properties: this.convertPropertiesToGeminiFormat(schema.properties || {}),
              required: schema.required as string[] || []
            }
          }] as [{
            name: string;
            description: string;
            parameters: {
              type: "OBJECT";
              properties: Record<string, any>;
              required?: string[];
            };
          }]
        };
      });

      log.info({ toolCount: geminiTools.length }, 'Gemini function calling tools from enabled servers prepared successfully');
      return geminiTools;
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      log.error({ error: errorMsg }, 'Failed to fetch Gemini function calling tools');
      return [];
    }
  }

  /**
   * OpenAI形式のプロパティをGemini形式に変換するヘルパー関数
   * Geminiは型の表記が異なるため変換が必要
   */
  private convertPropertiesToGeminiFormat(properties: Record<string, any>): Record<string, any> {
    const result: Record<string, any> = {};

    for (const [key, prop] of Object.entries(properties)) {
      // プロパティをコピー
      result[key] = { ...prop };

      // Geminiでサポートされていない属性を削除
      const unsupportedAttributes = [
        'additionalProperties',
        'default',
        'examples',
        'const',
        'multipleOf',
        'exclusiveMinimum',
        'exclusiveMaximum'
      ];

      for (const attr of unsupportedAttributes) {
        if (attr in result[key]) {
          delete result[key][attr];
        }
      }

      // 型の変換（小文字から大文字へ）
      if (prop.type) {
        // Geminiの型は大文字なので変換
        const typeMap: Record<string, string> = {
          'string': 'STRING',
          'number': 'NUMBER',
          'integer': 'INTEGER',
          'boolean': 'BOOLEAN',
          'array': 'ARRAY',
          'object': 'OBJECT'
        };

        result[key].type = typeMap[prop.type.toLowerCase()] || prop.type.toUpperCase();
      }

      // ネストされたobjectプロパティを再帰的に処理
      if (prop.type === 'object' && prop.properties) {
        result[key].properties = this.convertPropertiesToGeminiFormat(prop.properties);
      }

      // array型の場合、itemsも処理
      if (prop.type === 'array' && prop.items) {
        if (prop.items.type) {
          const itemsType = prop.items.type.toLowerCase();
          const typeMap: Record<string, string> = {
            'string': 'STRING',
            'number': 'NUMBER',
            'integer': 'INTEGER',
            'boolean': 'BOOLEAN',
            'array': 'ARRAY',
            'object': 'OBJECT'
          };
          result[key].items = { ...prop.items, type: typeMap[itemsType] || prop.items.type.toUpperCase() };

          // items内の非対応プロパティも削除
          for (const attr of unsupportedAttributes) {
            if (attr in result[key].items) {
              delete result[key].items[attr];
            }
          }
        }

        // itemsがobjectでpropertiesを持つ場合、再帰的に処理
        if (prop.items.type === 'object' && prop.items.properties) {
          result[key].items.properties = this.convertPropertiesToGeminiFormat(prop.items.properties);
        }
      }
    }

    return result;
  }

  /**
   * Gemini Function Callingの呼び出し結果を実際のMCPツール呼び出しに変換する
   * @param functionCall Geminiから返された関数呼び出し情報
   * @returns MCP Tool呼び出し結果
   */
  async executeGeminiFunctionCall(functionCall: {
    name: string;
    arguments: Record<string, any>;
  }): Promise<McpToolCallResponse> {
    const log = logger.child({ functionCall: functionCall.name });
    log.info('Executing function call from Gemini');

    try {
      // 関数名からサーバー名とツール名を抽出
      // 二重アンダースコア(__）を区切りとして使用
      const nameParts = functionCall.name.split('__');
      const serverName = nameParts[0];
      const toolName = nameParts.length > 1 ? nameParts[1] : functionCall.name;

      // 引数を取得（Geminiの場合はすでにオブジェクト形式）
      const toolArguments = functionCall.arguments || {};

      // MCPツールを呼び出す
      log.info({ serverName, toolName }, 'Calling MCP tool from Gemini function call');
      return await this.callTool(serverName, toolName, toolArguments);
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      log.error({ error: errorMsg }, 'Gemini function call execution failed');
      throw error;
    }
  }



  // Fetch resources list
  async fetchResourcesList(serverName: string): Promise<McpResource[]> {
    const log = logger.child({ serverName });
    log.debug('Fetching resources list');

    try {
      const connection = this.connections.get(serverName);
      if (!connection || !connection.client) {
        log.warn('Connection not available, returning empty resources list');
        return [];
      }

      const response = await connection.client.request(
        { method: "resources/list" },
        ListResourcesResultSchema
      );

      const resources = (response?.resources || []).map(resource => ({
        ...resource,
        description: resource.description || '', // Ensure description is always a string
      }));

      log.info({ resourceCount: resources.length }, 'Resources fetched successfully');
      return resources;
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Failed to fetch resources list');
      return [];
    }
  }

  // Fetch resource templates list
  async fetchResourceTemplatesList(serverName: string): Promise<McpResourceTemplate[]> {
    const log = logger.child({ serverName });
    log.debug('Fetching resource templates list');

    try {
      const connection = this.connections.get(serverName);
      if (!connection || !connection.client) {
        log.warn('Connection not available, returning empty templates list');
        return [];
      }

      try {
        // Improved error handling
        const response = await connection.client.request(
          { method: "resources/templates/list" },
          ListResourceTemplatesResultSchema
        );

        const templates = (response?.resourceTemplates || []).map(template => ({
          ...template,
          kind: 'template',
          description: template.description || '',
        }));

        log.info({ templateCount: templates.length }, 'Resource templates fetched successfully');
        return templates;
      } catch (error) {
        // Return empty array in case of error
        log.warn('Resource template feature is not supported by this server');
        return [];
      }
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Failed to fetch resource templates');
      return [];
    }
  }

  // Toggle server disabled status
  async toggleServerDisabled(serverName: string, disabled: boolean): Promise<boolean> {
    const log = logger.child({ serverName });
    log.info({ disabled }, 'Toggling server disabled status');

    try {
      // Load settings file
      const settings = await loadMcpSettings();

      if (!settings.mcpServers || !settings.mcpServers[serverName]) {
        const errorMsg = `Server "${serverName}" not found in settings`;
        log.error({ error: errorMsg }, 'Toggle server disabled failed');
        throw new Error(errorMsg);
      }

      // Update server configuration
      settings.mcpServers[serverName].disabled = disabled;

      // Save settings file
      await saveMcpSettings(settings);

      // Update connection information
      const connection = this.connections.get(serverName);
      if (connection) {
        connection.server.disabled = disabled;

        // Update feature lists if connected
        if (connection.server.status === 'connected') {
          connection.server.tools = await this.fetchToolsList(serverName);
          connection.server.resources = await this.fetchResourcesList(serverName);
          connection.server.resourceTemplates = await this.fetchResourceTemplatesList(serverName);
        }
      }

      log.info('Server disabled status toggled successfully');
      return true;
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Failed to toggle server status');
      throw error;
    }
  }

  // Toggle tool autoApprove setting
  async toggleToolAutoApprove(serverName: string, toolName: string, shouldAllow: boolean): Promise<boolean> {
    const log = logger.child({ serverName, toolName });
    log.info({ shouldAllow }, 'Toggling tool auto-approve setting');

    try {
      // Load settings file
      const settings = await loadMcpSettings();

      if (!settings.mcpServers || !settings.mcpServers[serverName]) {
        const errorMsg = `Server "${serverName}" not found in settings`;
        log.error({ error: errorMsg }, 'Toggle tool auto-approve failed');
        throw new Error(errorMsg);
      }

      // Initialize autoApprove list
      if (!settings.mcpServers[serverName].autoApprove) {
        settings.mcpServers[serverName].autoApprove = [];
      }

      const autoApprove = settings.mcpServers[serverName].autoApprove!;
      const toolIndex = autoApprove.indexOf(toolName);

      // Update list
      if (shouldAllow && toolIndex === -1) {
        autoApprove.push(toolName);
      } else if (!shouldAllow && toolIndex !== -1) {
        autoApprove.splice(toolIndex, 1);
      }

      // Save settings file
      await saveMcpSettings(settings);

      // Update tools list
      const connection = this.connections.get(serverName);
      if (connection) {
        connection.server.tools = await this.fetchToolsList(serverName);
      }

      log.info('Tool auto-approve setting toggled successfully');
      return true;
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Failed to toggle tool auto-approve');
      throw error;
    }
  }

  // Delete server
  async deleteServer(serverName: string): Promise<boolean> {
    const log = logger.child({ serverName });
    log.info('Deleting server');

    try {
      // Delete connection
      await this.deleteConnection(serverName);

      // Update settings file
      const settings = await loadMcpSettings();

      if (settings.mcpServers && settings.mcpServers[serverName]) {
        delete settings.mcpServers[serverName];
        await saveMcpSettings(settings);
      }

      log.info('Server deleted successfully');
      return true;
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Failed to delete server');
      throw error;
    }
  }

  // Update server timeout setting
  async updateServerTimeout(serverName: string, timeout: number): Promise<boolean> {
    const log = logger.child({ serverName });
    log.info({ timeout }, 'Updating server timeout setting');

    try {
      // Validate timeout value
      if (typeof timeout !== 'number' || timeout < config.minMcpTimeout) {
        const errorMsg = `Invalid timeout value: ${timeout}. Minimum value is ${config.minMcpTimeout} seconds.`;
        log.error({ error: errorMsg }, 'Update timeout failed');
        throw new Error(errorMsg);
      }

      // Update settings file
      const settings = await loadMcpSettings();

      if (!settings.mcpServers || !settings.mcpServers[serverName]) {
        const errorMsg = `Server "${serverName}" not found in settings`;
        log.error({ error: errorMsg }, 'Update timeout failed');
        throw new Error(errorMsg);
      }

      // Update timeout value
      settings.mcpServers[serverName].timeout = timeout;

      // Save settings file
      await saveMcpSettings(settings);

      // Update server connections
      await this.updateServerConnections(settings.mcpServers);

      log.info('Server timeout setting updated successfully');
      return true;
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Failed to update timeout setting');
      throw error;
    }
  }


  /**
   * callTool - Call a tool on the specified server
   * @param serverName 
   * @param toolName 
   * @param toolArguments 
   * @returns 
   */
  async callTool(serverName: string, toolName: string, toolArguments: Record<string, any>): Promise<McpToolCallResponse> {
    const log = logger.child({ serverName, toolName });
    log.info({ arguments: toolArguments }, 'Calling MCP tool');

    // Validate tool name and arguments
    const connection = this.connections.get(serverName);
    if (!connection) {
      const errorMsg = `Connection to server "${serverName}" not found`;
      log.error({ error: errorMsg }, 'Tool call failed');
      throw new Error(errorMsg);
    }

    // Check if server is disabled
    if (connection.server.disabled) {
      const errorMsg = `Server "${serverName}" is disabled and cannot be used`;
      log.error({ error: errorMsg }, 'Tool call failed');
      throw new Error(errorMsg);
    }

    // Check if client is initialized
    if (!connection.client) {
      const errorMsg = `Client for server "${serverName}" is not initialized`;
      log.error({ error: errorMsg }, 'Tool call failed');
      throw new Error(errorMsg);
    }

    // Get timeout value
    let timeout = secondsToMs(config.defaultMcpTimeout);
    try {
      const serverConfig = JSON.parse(connection.server.config) as ServerConfig;
      if (serverConfig.timeout && typeof serverConfig.timeout === 'number') {
        timeout = secondsToMs(serverConfig.timeout);
      }
    } catch (error) {
      log.warn({ error: error instanceof Error ? error.message : String(error) }, 'Error parsing timeout setting');
    }

    try {
      log.debug({ timeout }, 'Sending tool call request');
      const response = await connection.client.request(
        {
          method: "tools/call",
          params: {
            name: toolName,
            arguments: toolArguments,
          },
        },
        CallToolResultSchema,
        {
          timeout,
        }
      );

      log.info('Tool call completed successfully');
      // Transform the response to match McpToolCallResponse interface
      return { result: response };
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Tool call failed');
      throw error;
    }
  }

  // Read resource
  async readResource(serverName: string, uri: string): Promise<McpResourceResponse> {
    const log = logger.child({ serverName, resourceUri: uri });
    log.info('Reading resource');

    const connection = this.connections.get(serverName);
    if (!connection) {
      const errorMsg = `Connection to server "${serverName}" not found`;
      log.error({ error: errorMsg }, 'Resource read failed');
      throw new Error(errorMsg);
    }

    if (connection.server.disabled) {
      const errorMsg = `Server "${serverName}" is disabled`;
      log.error({ error: errorMsg }, 'Resource read failed');
      throw new Error(errorMsg);
    }

    if (!connection.client) {
      const errorMsg = `Client for server "${serverName}" is not initialized`;
      log.error({ error: errorMsg }, 'Resource read failed');
      throw new Error(errorMsg);
    }

    try {
      const response = await connection.client.request(
        {
          method: "resources/read",
          params: {
            uri,
          },
        },
        ReadResourceResultSchema
      );

      log.info({ mediaType: response.mediaType }, 'Resource read successfully');
      return {
        content: String(response.content),
        mediaType: response.mediaType as string
      };
    } catch (error) {
      log.error({ error: error instanceof Error ? error.message : String(error) }, 'Resource read failed');
      throw error;
    }
  }

  // Get all servers
  getAllServers(): McpServer[] {
    const servers = Array.from(this.connections.values()).map(conn => conn.server);
    logger.debug({ serverCount: servers.length }, 'Getting all servers');
    return servers;
  }

  // Close all connections
  async closeAll(): Promise<void> {
    logger.info({ connectionCount: this.connections.size }, 'Closing all connections');

    for (const name of this.connections.keys()) {
      await this.deleteConnection(name);
    }

    // Close all file watchers
    for (const watcher of this.fileWatchers.values()) {
      watcher.close();
    }
    this.fileWatchers.clear();

    logger.info('All connections closed successfully');
  }

  // Get internal state (mainly for testing)
  getState(): ServerManagerState {
    return {
      connections: this.connections,
      fileWatchers: this.fileWatchers,
      isConnecting: this.isConnecting
    };
  }
}