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

import pino, { LoggerOptions, TransportTargetOptions } from 'pino';

// Get log level from environment variables (default is 'info')
const LOG_LEVEL = process.env.LOG_LEVEL || 'info';
const NODE_ENV = process.env.NODE_ENV || 'development';

// Use readable format in development, JSON format in production
const isPrettyPrint = NODE_ENV === 'development';

// Logger configuration
const loggerOptions: LoggerOptions = {
  level: LOG_LEVEL,
  transport: isPrettyPrint
    ? {
        target: 'pino-pretty',
        options: {
          colorize: true,
          levelFirst: true,
          translateTime: 'SYS:standard',
          ignore: 'pid,hostname',
        },
      } as TransportTargetOptions
    : undefined,
  // Base information to easily identify the server
  base: {
    app: 'node-mcp-bridge',
    env: NODE_ENV,
  },
};

// Create logger instance
const logger = pino(loggerOptions);

export default logger;