#!/usr/bin/env node
/**
 * Office MCP Server - Main entry point
 * 
 * A Model Context Protocol server that provides access to
 * Microsoft 365 services through the Microsoft Graph API.
 */

// Load environment variables from .env file
// Use absolute path to ensure it loads regardless of working directory
const path = require('path');
const repoRootEnv = path.join(__dirname, '..', '..', '.env');
require('dotenv').config({ path: repoRootEnv });
require('dotenv').config({ path: path.join(__dirname, '.env') });

const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const { ListToolsRequestSchema, CallToolRequestSchema, McpError, ErrorCode } = require("@modelcontextprotocol/sdk/types.js");
const config = require('./config');
const { auditLog } = require('./utils/audit');
const { enforceOfficeToolPolicy } = require('./policy');

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const teamsTools = require('./teams');
const { notificationTools } = require('./notifications');
const { plannerTools } = require('./planner');
const { filesTools } = require('./files');
const { searchTools } = require('./search');
const { contactsTools } = require('./contacts');
const { todoTools } = require('./todo');
const { groupsTools } = require('./groups');
const { directoryTools } = require('./directory');
const { runWithAccessToken } = require('./auth/token-manager');

function extractBearerToken(req) {
  const authHeader = req.headers?.authorization || req.headers?.Authorization;
  if (typeof authHeader !== 'string') {
    return null;
  }
  const match = authHeader.match(/^Bearer\s+(.+)$/i);
  return match?.[1]?.trim() || null;
}

// Log startup information
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);
console.error(`Client ID: ${config.AUTH_CONFIG.clientId ? config.AUTH_CONFIG.clientId.substring(0, 8) + '...' : 'NOT SET'}`);
console.error(`Token path: ${config.AUTH_CONFIG.tokenStorePath}`);
console.error(`Token exists: ${require('fs').existsSync(config.AUTH_CONFIG.tokenStorePath)}`);

// Combine all tools
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...teamsTools,
  ...notificationTools,
  ...plannerTools,
  ...filesTools,
  ...searchTools,
  ...contactsTools,
  ...todoTools,
  ...groupsTools,
  ...directoryTools
];

/** Each SSE client needs its own Server — the SDK allows one transport per instance. */
function createMcpServer() {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => {
    console.error(`TOOLS LIST REQUEST — ${TOOLS.length} tools`);
    return {
      tools: TOOLS.map(tool => ({
        name: tool.name,
        description: tool.description,
        inputSchema: tool.inputSchema
      }))
    };
  });

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args = {} } = request.params;
    console.error(`TOOL CALL: ${name}`);
    auditLog(name, args);

    const tool = TOOLS.find(t => t.name === name);
    if (!tool) {
      throw new McpError(ErrorCode.MethodNotFound, `Tool not found: ${name}`);
    }
    const policyError = await enforceOfficeToolPolicy(name, args);
    if (policyError) {
      return { content: [{ type: 'text', text: policyError }], isError: true };
    }
    return await tool.handler(args);
  });

  return server;
}

// Graceful shutdown handlers
let isShuttingDown = false;

function gracefulShutdown(signal) {
  if (isShuttingDown) return;
  isShuttingDown = true;

  console.error(`[SHUTDOWN] ${signal} received, shutting down gracefully`);

  // Give pending operations time to complete
  setTimeout(() => {
    console.error('[SHUTDOWN] Exiting');
    process.exit(0);
  }, 1000);
}

process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));
process.on('SIGINT', () => gracefulShutdown('SIGINT'));

function removeSession(sessions, sessionId) {
  const session = sessions.get(sessionId);
  if (!session) return;
  sessions.delete(sessionId);
  void session.server.close().catch((error) => {
    console.error(`[SSE] Failed to close session ${sessionId}: ${error.message}`);
  });
}

// Start the server
if (config.TRANSPORT_TYPE === 'http') {
  const { SSEServerTransport } = require("@modelcontextprotocol/sdk/server/sse.js");
  const express = require('express');
  const app = express();
  app.use(express.json());

  const sessions = new Map();

  app.get("/sse", async (req, res) => {
    const bearerToken = extractBearerToken(req);
    const transport = new SSEServerTransport("/message", res);
    const server = createMcpServer();
    const sessionId = transport.sessionId;

    sessions.set(sessionId, { server, transport, bearerToken });
    console.error(`[SSE] Client connected (session ${sessionId}, active=${sessions.size})`);

    const cleanup = () => {
      removeSession(sessions, sessionId);
      console.error(`[SSE] Client disconnected (session ${sessionId}, active=${sessions.size})`);
    };
    res.on("close", cleanup);

    try {
      await runWithAccessToken(bearerToken, async () => {
        await server.connect(transport);
      });
    } catch (error) {
      cleanup();
      console.error(`[SSE] Connection error (session ${sessionId}): ${error.message}`);
      if (!res.headersSent) {
        res.status(500).end("SSE connection failed");
      }
    }
  });

  app.post("/message", async (req, res) => {
    const sessionId = req.query?.sessionId;
    const session =
      (typeof sessionId === "string" ? sessions.get(sessionId) : undefined) ||
      (sessions.size === 1 ? sessions.values().next().value : undefined);

    if (!session) {
      res.status(503).json({ error: "No SSE connection for session" });
      return;
    }

    const bearerToken = extractBearerToken(req) || session.bearerToken;
    try {
      await runWithAccessToken(bearerToken, async () => {
        await session.transport.handlePostMessage(req, res, req.body);
      });
    } catch (error) {
      console.error(`[SSE] Message error (session ${session.transport.sessionId}): ${error.message}`);
      if (!res.headersSent) {
        res.status(500).json({ error: "Failed to handle MCP message" });
      }
    }
  });

  app.listen(config.HTTP_PORT, config.HTTP_HOST, () => {
    console.error(`${config.SERVER_NAME} HTTP/SSE on ${config.HTTP_HOST}:${config.HTTP_PORT}`);
  });
} else {
  const transport = new StdioServerTransport();
  const server = createMcpServer();
  server.connect(transport)
    .then(() => console.error(`${config.SERVER_NAME} connected and listening`))
    .catch(error => {
      console.error(`Connection error: ${error.message}`);
      process.exit(1);
    });
}
