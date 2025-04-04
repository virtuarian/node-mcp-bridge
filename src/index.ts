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

/**
 * Node MCP Bridge
 * Middleware responsible for coordination between callers and Model Context Protocol (MCP) servers
 */

import express from 'express';
import path from 'path';
import chokidar from 'chokidar';
import { fileURLToPath } from 'url';
import { config, loadMcpSettings, validateServerConfig, saveMcpSettings } from './config.js';
import { MCPServerManager } from './server-manager.js';
import logger from './logger.js';
import pinoHttpModule from 'pino-http';
const pinoHttp = pinoHttpModule.default;
import { LowDBSessionManager as SessionManager } from './lowdb-session-manager.js';

// Alternative to __dirname (for ES modules)
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);


// Main function
async function main() {
  try {
    logger.info('MCP Bridge started');

    // Create server management and session management instances
    const serverManager = new MCPServerManager();

    // Load initial server settings
    const settings = await loadMcpSettings();
    await serverManager.updateServerConnections(settings.mcpServers || {});

    // Create session manager instance
    const sessionManager = new SessionManager();

    // Express application
    const app = express();

    // Add HTTP request logging middleware
    app.use(pinoHttp({
      logger,
      // Exclude health check endpoint from logs to reduce noise
      autoLogging: {
        ignore: (req) => req.url === '/health'
      }
    }));

    app.use(express.json());

    // Configuring static file serving
    const adminPath = path.join(__dirname, 'admin');
    app.use('/admin/static', express.static(adminPath));

    // Admin UI endpoint
    app.get('/admin', (req, res) => {
      res.sendFile(path.join(adminPath, 'index.html'));
    });

    // Health check endpoint
    app.get('/health', (req, res) => {
      res.status(200).json({ status: 'ok' });
    });


    // Tools list endpoint
    app.get('/tools', async (req, res) => {
      try {
        // Get tools list from servers
        const serverTools = await serverManager.fetchAllToolsList();
        const tools = serverTools.map(tool => ({
          serverName: tool.serverName,
          name: tool.name,
          description: tool.description || `${tool.name} tool`,
          parameters: tool.schema || {
            type: 'object',
            properties: {},
            required: []
          },
          autoApprove: tool.autoApprove
        }));

        res.status(200).json({ tools });
      }
      catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);
        logger.error({ error: errorMsg }, 'Error fetching tools list');
        res.status(500).json({ error: errorMsg });
      }
    });

    // === MCP Command API ===

    // セッション作成エンドポイント
    // app.post('/sessions', (req, res) => {
    //   const sessionId = sessionManager.createSession();
    //   res.status(200).json({ sessionId });
    // });

    // ツール承認エンドポイント
    app.post('/tools/call/:sessionId/approve', async (req, res) => {
      const { sessionId } = req.params;
      // const { serverName, toolName } = req.body;
      const { toolName, serverName, arguments: toolArguments } = req.body;

      if (!sessionId || !serverName || !toolName) {
        return res.status(400).json({
          error: 'sessionId, serverName, and toolName are required'
        });
      }

      if (!await sessionManager.hasSession(sessionId)) {
        return res.status(404).json({ error: 'Session not found or expired' });
      }

      // Check if the tool is already approved
      const approved = await sessionManager.approveToolForSession(
        sessionId,
        serverName,
        toolName
      );

      if (approved) {
        // Fetch tools list from server
        logger.info({ sessionId, toolName, serverName }, 'Calling approved tool');
        const result = await serverManager.callTool(serverName, toolName, toolArguments);

        res.status(200).json(result);
      } else {
        res.status(500).json({ error: 'Failed to approve tool' });
      }
    });

    // セッションベースのツール呼び出しエンドポイント
    app.post('/tools/call/:sessionId', async (req, res) => {
      try {
        const { sessionId } = req.params;
        const { toolName, serverName, arguments: toolArguments } = req.body;

        logger.info({ sessionId, toolName, serverName, arguments: toolArguments }, 'Calling tool with session');

        if (!sessionId) {
          return res.status(400).json({ error: 'sessionId is required' });
        }

        if (!toolName) {
          return res.status(400).json({ error: 'toolName is required' });
        }

        if (!await sessionManager.hasSession(sessionId)) {
          return res.status(404).json({ error: 'Session not found or expired' });
        }

        logger.info({ sessionId, toolName, serverName }, 'Fetching tools list for server');

        // ツール情報を取得
        const tools = await serverManager.fetchToolsList(serverName);
        const tool = tools.find(t => t.name === toolName);

        if (!tool) {
          return res.status(404).json({
            error: `Tool "${toolName}" not found on server "${serverName}"`
          });
        }

        logger.info({ sessionId, toolName, serverName, tool }, 'Checking tool approval status');

        // 自動承認がOFFで、セッションでの承認もない場合は承認要求
        if (!tool.autoApprove &&
          !await sessionManager.isToolApproved(sessionId, serverName, toolName)) {
          return res.status(403).json({
            error: 'Tool requires approval',
            approvalRequired: true,
            serverName,
            toolName,
            toolDescription: tool.description || `${toolName} tool`,
            sessionId
          });
        }

        // 承認済みまたは自動承認の場合はツールを実行
        logger.info({ sessionId, toolName, serverName }, 'Calling approved tool');
        const result = await serverManager.callTool(serverName, toolName, toolArguments);

        res.status(200).json(result);
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);
        logger.error({
          error: errorMsg,
          sessionId: req.params.sessionId
        }, 'Tool call error');
        res.status(500).json({ error: errorMsg });
      }
    });

    // Tool call endpoint
    app.post('/tools/call', async (req, res) => {
      try {
        const { toolName, serverName, arguments: toolArguments } = req.body;

        logger.info({ toolName, serverName, arguments: toolArguments }, 'Calling tool without session');

        if (!toolName) {
          return res.status(400).json({ error: 'toolName is required' });
        }

        logger.info({ toolName, serverName, arguments: toolArguments }, 'Calling tool');
        const result = await serverManager.callTool(serverName, toolName, toolArguments);

        res.status(200).json(result);
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);
        logger.error({ error: errorMsg }, 'Tool call error');
        res.status(500).json({ error: errorMsg });
      }
    });

    // Endpoint to get a list of tools for a specific server
    app.get('/admin/servers/:serverName/tools', async (req, res) => {
      try {
        const tools = await serverManager.fetchToolsList(req.params.serverName);
        res.status(200).json({ tools });
      } catch (error) {
        res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    // === Admin API ===

    // Server list endpoint
    app.get('/admin/servers', (req, res) => {
      const servers = serverManager.getAllServers();
      res.status(200).json(servers);
    });

    // Server restart endpoint
    app.post('/admin/servers/:serverName/restart', async (req, res) => {
      try {
        await serverManager.restartConnection(req.params.serverName);
        res.status(200).json({ status: 'restarted' });
      } catch (error) {
        res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    // Server enable/disable toggle endpoint
    app.put('/admin/servers/:serverName/toggleDisabled', async (req, res) => {
      try {
        const { disabled } = req.body;
        if (disabled === undefined) {
          return res.status(400).json({ error: 'disabled parameter is required' });
        }

        await serverManager.toggleServerDisabled(req.params.serverName, disabled);
        res.status(200).json({ status: 'updated' });
      } catch (error) {
        res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    // Tool auto-approve toggle endpoint
    app.put('/admin/servers/:serverName/tools/:toolName/toggleAutoApprove', async (req, res) => {
      try {
        const { shouldAllow } = req.body;
        if (shouldAllow === undefined) {
          return res.status(400).json({ error: 'shouldAllow parameter is required' });
        }

        await serverManager.toggleToolAutoApprove(
          req.params.serverName,
          req.params.toolName,
          shouldAllow
        );

        res.status(200).json({ status: 'updated' });
      } catch (error) {
        res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    // Server timeout setting update endpoint
    app.put('/admin/servers/:serverName/timeout', async (req, res) => {
      try {
        const { timeout } = req.body;
        if (!timeout || typeof timeout !== 'number') {
          return res.status(400).json({ error: 'Valid timeout parameter is required' });
        }

        await serverManager.updateServerTimeout(req.params.serverName, timeout);
        res.status(200).json({ status: 'updated' });
      } catch (error) {
        res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    // Server deletion endpoint
    app.delete('/admin/servers/:serverName', async (req, res) => {
      try {
        await serverManager.deleteServer(req.params.serverName);
        res.status(204).end();
      } catch (error) {
        res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    // Server add/update endpoint
    app.put('/admin/servers/:serverName', async (req, res) => {
      try {
        const serverConfig = req.body;
        if (!validateServerConfig(serverConfig)) {
          return res.status(400).json({ error: 'Invalid server configuration' });
        }

        // Load settings file
        const settings = await loadMcpSettings();
        if (!settings.mcpServers) {
          settings.mcpServers = {};
        }

        // Update settings
        settings.mcpServers[req.params.serverName] = serverConfig;

        // Save settings
        await saveMcpSettings(settings);

        // Update server connections
        await serverManager.updateServerConnections(settings.mcpServers);

        res.status(200).json({ status: 'updated' });
      } catch (error) {
        res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
      }
    });

    // Start server
    const port = config.port;
    app.listen(port, () => {
      logger.info({ port }, `Start Node MCP Bridge: http://localhost:${port}`);
    });

    // Settings file monitoring (automatically update server settings when changes occur)
    const settingsDir = path.dirname(config.settingsPath);
    const settingsFile = path.basename(config.settingsPath);

    const watcher = chokidar.watch(config.settingsPath, {
      persistent: true,
      ignoreInitial: true,
      awaitWriteFinish: true,
    });

    watcher.on('change', async () => {
      logger.info('Configuration file change detected. Updating server settings...');
      try {
        const newSettings = await loadMcpSettings();
        await serverManager.updateServerConnections(newSettings.mcpServers || {});
        logger.info('Server settings updated successfully');
      } catch (error) {
        logger.error({ error }, 'Settings update error');
      }
    });

    // Process termination handling
    process.on('SIGINT', async () => {
      logger.info('Shutting down...');
      watcher.close();
      await serverManager.closeAll();
      process.exit(0);
    });

    process.on('SIGTERM', async () => {
      logger.info('Shutting down...');
      watcher.close();
      await serverManager.closeAll();
      process.exit(0);
    });

  } catch (error) {
    logger.fatal({ error }, 'Startup error');
    process.exit(1);
  }
}

// Application startup
main().catch(err => {
  logger.fatal({ error: err }, 'Application Start Error');
  process.exit(1);
});