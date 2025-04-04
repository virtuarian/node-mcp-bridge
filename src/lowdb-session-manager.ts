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

import path, { join, dirname } from 'path';
import { Low } from 'lowdb';
import { JSONFile } from 'lowdb/node';
import { v4 as uuidv4 } from 'uuid';
import fs from 'fs';
import logger from './logger.js';
import { config, loadMcpSettings } from './config.js';
import { McpSettings } from './types.js';

// Define types
interface SessionState {
    id: string;
    approvedTools: Record<string, string[]>;
    createdAt: string;
    lastActive: string;
}

interface DbData {
    sessions: Record<string, SessionState>;
}

export class LowDBSessionManager {
    private db!: Low<DbData>;
    private readonly defaultSessionExpiry: number = config.defaultSessionTimeout * 60 * 1000; // デフォルトは設定値
    private initialized: boolean = false;
    private mcpSettings: McpSettings | null = null;

    constructor(private dbPath: string = config.sessionDbPath) {
        this.init();
    }

    private async init() {
        try {
            // ログに絶対パスを出力して確認しやすくする
            const absoluteDbPath = path.resolve(this.dbPath);
            logger.info({ dbPath: this.dbPath, absoluteDbPath }, 'Initializing session storage');

            // Ensure directory exists
            const dir = dirname(this.dbPath);
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
                logger.info({ directory: dir }, 'Created session data directory');
            }

            // Setup adapter and db
            const adapter = new JSONFile<DbData>(this.dbPath);
            this.db = new Low<DbData>(adapter, { sessions: {} });

            // Load initial data
            await this.db.read();

            // Ensure sessions object exists
            if (!this.db.data.sessions) {
                this.db.data.sessions = {};
            }

            // MCP設定の読み込み
            this.mcpSettings = await loadMcpSettings();

            this.initialized = true;
            logger.info({ dbPath: this.dbPath }, 'LowDB Session Manager initialized');

            // Schedule settings reload
            this.scheduleSettingsReload();

            // Start cleanup process
            this.scheduleCleanup();
        } catch (error) {
            logger.error({ error, dbPath: this.dbPath }, 'Failed to initialize LowDB session storage');
        }
    }

    private scheduleCleanup() {
        // Run cleanup every hour
        setInterval(() => this.cleanupExpiredSessions(), 60 * 60 * 1000);
    }

    // サーバーのセッションタイムアウト（ミリ秒）を取得する
    private getSessionExpiryForServer(serverName: string): number {
        if (!this.mcpSettings || !this.mcpSettings.mcpServers || !this.mcpSettings.mcpServers[serverName]) {
            return this.defaultSessionExpiry;
        }

        const serverConfig = this.mcpSettings.mcpServers[serverName];

        // 0 = 無制限
        if (serverConfig.sessionTimeout === 0) {
            return Number.MAX_SAFE_INTEGER;
        }

        return (serverConfig.sessionTimeout || config.defaultSessionTimeout) * 60 * 1000;
    }

    // Create a new session with a unique ID
    async createSession(sessionId:string): Promise<SessionState> {
        if (!this.initialized) await this.init();

        // const sessionId = uuidv4();
        const session: SessionState = {
            id: sessionId,
            approvedTools: {},
            createdAt: new Date().toISOString(),
            lastActive: new Date().toISOString()
        };

        this.db.data.sessions[sessionId] = session;
        await this.db.write();

        logger.info({ sessionId }, 'Created new session');
        return session;
    }

    async hasSession(sessionId: string): Promise<boolean> {
        if (!this.initialized) await this.init();
        if (!sessionId) return false;

        try {
            let session = this.db.data.sessions[sessionId];
            if (!session) {
                session = await this.createSession(sessionId);
            }

            logger.debug({ sessionId }, 'Session found');

            // Update last active time
            let maxExpiry = this.defaultSessionExpiry;
            const lastActive = new Date(session.lastActive).getTime();
            const now = Date.now();

            // Check if session expired
            if (session.approvedTools) {
                for (const serverName of Object.keys(session.approvedTools)) {
                    const serverExpiry = this.getSessionExpiryForServer(serverName);
                    maxExpiry = Math.max(maxExpiry, serverExpiry);

                    // 0 : unlimited
                    if (maxExpiry === Number.MAX_SAFE_INTEGER) {
                        return true;
                    }
                }
            }

            // Check if session expired based on last active time
            if (now - lastActive > maxExpiry) {
                logger.info({ sessionId }, 'Session expired');
                delete this.db.data.sessions[sessionId];
                await this.db.write();
                return false;
            }

            return true;
        } catch (error) {
            logger.error({ error, sessionId }, 'Error checking session');
            return false;
        }
    }

    // Get session state by ID
    async getSession(sessionId: string): Promise<SessionState | null> {
        if (!this.initialized) await this.init();
        if (!sessionId) return null;

        try {
            const session = this.db.data.sessions[sessionId];
            if (!session) {
                logger.debug({ sessionId }, 'Session not found when getting');
                return null;
            }

            logger.debug({ sessionId }, 'Session found when getting');

            // Check if session expired
            const lastActive = new Date(session.lastActive).getTime();
            const now = Date.now();
            if (now - lastActive > this.defaultSessionExpiry) {
                logger.info({ sessionId }, 'Session expired when getting');
                delete this.db.data.sessions[sessionId];
                await this.db.write();
                return null;
            }

            // Update last active time
            session.lastActive = new Date().toISOString();
            await this.db.write();

            return session;
        } catch (error) {
            logger.error({ error, sessionId }, 'Error getting session');
            return null;
        }
    }

    // Check if a tool is approved for a session
    async isToolApproved(sessionId: string, serverName: string, toolName: string): Promise<boolean> {
        if (!this.initialized) await this.init();

        try {

            const session = await this.getSession(sessionId);

            if (!session) return false;

            const isApproved = session.approvedTools[serverName]?.includes(toolName) || false;


            logger.debug({ sessionId, serverName, toolName, isApproved }, 'Checking tool approval status');
            return isApproved;
        } catch (error) {
            logger.error({ error, sessionId, serverName, toolName }, 'Error checking tool approval');
            return false;
        }
    }

    // Approve a tool for a session
    async approveToolForSession(sessionId: string, serverName: string, toolName: string): Promise<boolean> {
        if (!this.initialized) await this.init();

        try {
            const session = await this.getSession(sessionId);
            if (!session) {
                logger.warn({ sessionId }, 'Failed to approve tool: Session not found');
                return false;
            }

            // Initialize server tools array if it doesn't exist
            if (!session.approvedTools[serverName]) {
                session.approvedTools[serverName] = [];
            }

            // Add tool to approved list if not already there
            if (!session.approvedTools[serverName].includes(toolName)) {
                session.approvedTools[serverName].push(toolName);
                await this.db.write();
            }

            logger.info({ sessionId, serverName, toolName }, 'Tool approved for session');
            return true;
        } catch (error) {
            logger.error({ error, sessionId, serverName, toolName }, 'Error approving tool');
            return false;
        }
    }

    // Remove a tool from the approved list for a session
    async removeSession(sessionId: string): Promise<boolean> {
        if (!this.initialized) await this.init();

        try {
            if (this.db.data.sessions[sessionId]) {
                delete this.db.data.sessions[sessionId];
                await this.db.write();
                logger.info({ sessionId }, 'Session removed');
                return true;
            }
            return false;
        } catch (error) {
            logger.error({ error, sessionId }, 'Error removing session');
            return false;
        }
    }

    async cleanupExpiredSessions(): Promise<void> {
        if (!this.initialized) await this.init();

        try {
            const now = Date.now();
            let expiredCount = 0;

            Object.keys(this.db.data.sessions).forEach(sessionId => {
                const session = this.db.data.sessions[sessionId];
                const lastActive = new Date(session.lastActive).getTime();

                if (now - lastActive > this.defaultSessionExpiry) {
                    delete this.db.data.sessions[sessionId];
                    expiredCount++;
                }
            });

            if (expiredCount > 0) {
                await this.db.write();
                logger.info({ count: expiredCount }, 'Expired sessions cleaned up');
            }
        } catch (error) {
            logger.error({ error }, 'Error cleaning up expired sessions');
        }
    }

    // 定期的に設定をリロードする
    private scheduleSettingsReload() {
        // 5分ごとに設定を再読み込み
        setInterval(async () => {
            try {
                this.mcpSettings = await loadMcpSettings();
                logger.debug('MCP settings reloaded for session manager');
            } catch (error) {
                logger.error({ error }, 'Failed to reload MCP settings');
            }
        }, 5 * 60 * 1000);
    }
}