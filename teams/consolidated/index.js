/**
 * Consolidated Teams Tools Export
 *
 * Each tool is operation-based. inputSchema lists all parameters so MCP clients
 * can pass chatId, teamId, etc. (not just operation).
 */

const handleTeamsMeeting = require('./teams_meeting');
const handleTeamsChannel = require('./teams_channel');
const handleTeamsChat = require('./teams_chat');
const { safeTool } = require('../../utils/errors');

const meetingToolSchema = {
  type: 'object',
  required: ['operation'],
  properties: {
    operation: {
      type: 'string',
      description: 'The operation to perform',
      enum: [
        'create', 'update', 'cancel', 'get', 'find_by_url',
        'list_transcripts', 'get_transcript', 'list_recordings',
        'get_recording', 'get_participants', 'get_insights',
      ],
    },
    meetingId: {
      type: 'string',
      description: 'Online meeting ID (get, update, cancel, transcripts, recordings, participants, insights)',
    },
    subject: { type: 'string', description: 'Meeting subject (create, update)' },
    startDateTime: {
      type: 'string',
      description: 'Start time ISO 8601 (create, update)',
    },
    endDateTime: {
      type: 'string',
      description: 'End time ISO 8601 (create, update)',
    },
    description: { type: 'string', description: 'Meeting description (create, update)' },
    participants: {
      type: 'array',
      items: { type: 'string' },
      description: 'Participant emails or user IDs (create)',
    },
    joinUrl: { type: 'string', description: 'Teams join URL (find_by_url)' },
    transcriptId: { type: 'string', description: 'Transcript ID (get_transcript)' },
    recordingId: { type: 'string', description: 'Recording ID (get_recording)' },
    format: {
      type: 'string',
      description: 'Transcript format, e.g. text/vtt (get_transcript)',
    },
    comment: { type: 'string', description: 'Cancellation comment (cancel)' },
  },
};

const channelToolSchema = {
  type: 'object',
  required: ['operation'],
  properties: {
    operation: {
      type: 'string',
      description: 'The operation to perform',
      enum: [
        'list', 'create', 'get', 'update', 'delete',
        'list_messages', 'get_message', 'create_message', 'reply_to_message',
        'list_members', 'add_member', 'remove_member', 'list_tabs',
      ],
    },
    teamId: {
      type: 'string',
      description: 'Microsoft Teams team ID (required for all channel operations except none)',
    },
    channelId: {
      type: 'string',
      description: 'Channel ID within the team (get, update, delete, messages, members, tabs)',
    },
    displayName: { type: 'string', description: 'Channel name (create, update)' },
    description: { type: 'string', description: 'Channel description (create, update)' },
    membershipType: {
      type: 'string',
      enum: ['standard', 'private'],
      description: 'Channel visibility (create)',
    },
    messageId: {
      type: 'string',
      description: 'Channel message ID (get_message, reply_to_message)',
    },
    content: {
      type: 'string',
      description: 'Message body HTML or text (create_message, reply_to_message)',
    },
    maxResults: {
      type: 'number',
      description: 'Max messages to return (list_messages, default 25)',
    },
    attachments: {
      type: 'array',
      items: { type: 'object' },
      description: 'Message attachments (create_message, reply_to_message)',
    },
    userId: { type: 'string', description: 'User ID (add_member)' },
    email: { type: 'string', description: 'User email (add_member)' },
    memberId: { type: 'string', description: 'Member ID (remove_member)' },
    roles: {
      type: 'array',
      items: { type: 'string' },
      description: 'Member roles (add_member)',
    },
  },
};

const chatToolSchema = {
  type: 'object',
  required: ['operation'],
  properties: {
    operation: {
      type: 'string',
      description:
        'The operation to perform. Use reply_to_message (not send_message) to quote-reply to a specific message in group/1:1 chats.',
      enum: [
        'list', 'create', 'get', 'update', 'delete',
        'list_messages', 'get_message', 'send_message', 'reply_to_message',
        'update_message', 'delete_message',
        'list_members', 'add_member', 'remove_member',
      ],
    },
    chatId: {
      type: 'string',
      description:
        'Teams chat ID from teams_chat list (required for list_messages, get, send_message, and most non-list operations)',
    },
    recipient: {
      type: 'string',
      description:
        'send_message only: person name, email, or group topic to message when chatId is unknown (an existing chat is looked up)',
    },
    messageId: {
      type: 'string',
      description: 'Chat message ID from list_messages (reply_to_message, get_message, update_message, delete_message)',
    },
    messageIds: {
      type: 'array',
      items: { type: 'string' },
      description: 'Optional multiple message IDs to quote (reply_to_message, max 10)',
    },
    content: {
      type: 'string',
      description: 'Message body (send_message, reply_to_message, update_message)',
    },
    maxResults: {
      type: 'number',
      description: 'Max chats or messages to return (list, list_messages; default 25/50)',
    },
    topic: { type: 'string', description: 'Group chat topic (create, update)' },
    members: {
      type: 'array',
      items: { type: 'string' },
      description: 'Member user IDs or emails (create)',
    },
    replyToId: {
      type: 'string',
      description: 'Deprecated for group chats — use reply_to_message instead. replyToId only threads in channels.',
    },
    attachments: {
      type: 'array',
      items: { type: 'object' },
      description: 'Message attachments (send_message)',
    },
    userId: { type: 'string', description: 'User ID (add_member)' },
    email: { type: 'string', description: 'User email (add_member)' },
    memberId: { type: 'string', description: 'Chat member ID (remove_member)' },
    roles: {
      type: 'array',
      items: { type: 'string' },
      description: 'Member roles (add_member)',
    },
  },
};

module.exports = [
  {
    name: 'teams_meeting',
    description:
      'Teams meeting operations: create, update, cancel, find, list transcripts, get recordings, and more',
    inputSchema: meetingToolSchema,
    handler: safeTool('teams_meeting', handleTeamsMeeting),
  },
  {
    name: 'teams_channel',
    description:
      'Teams channel operations: list, create, get, update, delete channels and manage messages, members, and tabs. Requires teamId; channel ops also need channelId.',
    inputSchema: channelToolSchema,
    handler: safeTool('teams_channel', handleTeamsChannel),
  },
  {
    name: 'teams_chat',
    description:
      'Teams chat operations: list, create, get, update, delete chats and manage messages and members. Use list for chatId, list_messages for messageId, reply_to_message to quote-reply in group/1:1 chats.',
    inputSchema: chatToolSchema,
    handler: safeTool('teams_chat', handleTeamsChat),
  },
];
