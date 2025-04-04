import path from 'path';
import fs from 'fs/promises';
import { ServerConfig, McpSettings } from './types.js';

// Configuration
export const config = {
  port: process.env.PORT ? parseInt(process.env.PORT) : 3001,
  sessionTTL: 60 * 60, // 1 hour
  defaultMcpTimeout: 60, // seconds
  minMcpTimeout: 5, // minimum timeout in seconds
  defaultSessionTimeout: 180, // minutes (3 hours)
  settingsPath: process.env.SETTINGS_PATH || path.join(process.cwd(), 'config', 'mcp-settings.json'),
  sessionDbPath: process.env.SESSION_DB_PATH || path.join(process.cwd(), 'data', 'sessionsDb.json'),
};

// Validate MCP server configuration
export function validateServerConfig(serverConfig: any): serverConfig is ServerConfig {
  if (!serverConfig || typeof serverConfig !== 'object') return false;
  if (!serverConfig.command || typeof serverConfig.command !== 'string') return false;
  
  // args is optional, but must be an array if specified
  if (serverConfig.args && !Array.isArray(serverConfig.args)) return false;
  
  // env is optional, but must be an object if specified
  if (serverConfig.env && typeof serverConfig.env !== 'object') return false;
  
  // autoApprove is optional, but must be an array if specified
  if (serverConfig.autoApprove && !Array.isArray(serverConfig.autoApprove)) return false;
  
  // timeout is optional, but must be a number and not less than the minimum if specified
  if (serverConfig.timeout !== undefined) {
    if (typeof serverConfig.timeout !== 'number' || serverConfig.timeout < config.minMcpTimeout) return false;
  }
  
  // sessionTimeout is optional, but must be a number if specified
  if (serverConfig.sessionTimeout !== undefined && 
      (typeof serverConfig.sessionTimeout !== 'number' || serverConfig.sessionTimeout < 0)) {
    return false;
  }
  
  return true;
}

// Load MCP server settings
export async function loadMcpSettings(): Promise<McpSettings> {
  try {
    // Check if configuration directory exists, create if not
    const configDir = path.dirname(config.settingsPath);
    try {
      await fs.access(configDir);
    } catch (error) {
      await fs.mkdir(configDir, { recursive: true });
    }
    
    // Check if settings file exists, create default if not
    try {
      await fs.access(config.settingsPath);
      const content = await fs.readFile(config.settingsPath, 'utf-8');
      return JSON.parse(content) as McpSettings;
    } catch (error) {
      // Create default settings
      const defaultSettings: McpSettings = {
        mcpServers: {}
      };
      await fs.writeFile(config.settingsPath, JSON.stringify(defaultSettings, null, 2));
      return defaultSettings;
    }
  } catch (error) {
    console.error('Error loading settings file:', error);
    return { mcpServers: {} };
  }
}

// Save MCP server settings
export async function saveMcpSettings(settings: McpSettings): Promise<boolean> {
  try {
    await fs.writeFile(config.settingsPath, JSON.stringify(settings, null, 2));
    return true;
  } catch (error) {
    console.error('Error saving settings file:', error);
    return false;
  }
}

// Convert seconds to milliseconds
export function secondsToMs(seconds: number): number {
  return seconds * 1000;
}