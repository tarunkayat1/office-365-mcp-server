/**
 * Consolidated Teams Chat Module
 * 
 * Provides a unified interface for all Teams chat operations including:
 * - Listing and creating chats
 * - Getting and managing chat details
 * - Working with chat messages
 * - Managing chat members
 * - Handling message replies
 * - Working with chat attachments
 */

const { ensureAuthenticated } = require('../../auth');
const { callGraphAPI } = require('../../utils/graph-api');
const _config = require('../../config');

/**
 * Main handler for teams_chat operations
 * @param {Object} args - The operation arguments
 * @returns {Object} - MCP response
 */
async function handleTeamsChat(args) {
  // First extract the original operation parameter
  let { operation, ...params } = args;
  
  // Map legacy operation names to new ones for backward compatibility
  const operationMappings = {
    'get': 'get_message',      // Legacy chat_messages operation mapping
    'send': 'send_message',    // Legacy chat_messages operation mapping
    'update': 'update_message', // Legacy chat_messages operation mapping
    'delete': 'delete_message'  // Legacy chat_messages operation mapping
  };
  
  // Special handling for backward compatibility with chat_messages tool
  // If the client is using the old chat_messages tool pattern, adjust accordingly
  if (args.chatId && args.messageId && !operation) {
    // This looks like a legacy get_message call
    operation = 'get_message';
  } else if (args.chatId && args.content && !operation) {
    // This looks like a legacy send_message call
    operation = 'send_message';
  }
  
  // Map legacy operation names
  if (operation && operationMappings[operation]) {
    console.error(`Mapping legacy operation '${operation}' to '${operationMappings[operation]}'`);
    operation = operationMappings[operation];
  }
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, create, get, update, delete, list_messages, get_message, send_message, reply_to_message, update_message, delete_message, list_members, add_member, remove_member" 
      }]
    };
  }
  
  try {
    console.error(`teams_chat operation: ${operation}`);
    console.error('teams_chat params:', JSON.stringify(params));
    
    const accessToken = await ensureAuthenticated();

    // Small models garble or truncate long chat IDs and sometimes pass a name
    // instead. Repair the ID, or look the chat up, before calling Graph.
    if (operation !== 'list' && operation !== 'create') {
      const resolved = await resolveChatParam(accessToken, params);
      if (resolved.error) {
        return { content: [{ type: "text", text: resolved.error }] };
      }
      if (!params.chatId && operation === 'list_messages') {
        return await listLatestMessagesAcrossChats(accessToken, params.maxResults);
      }
    }

    switch (operation) {
      case 'list':
        return await listChats(accessToken, params);
      case 'create':
        return await createChat(accessToken, params);
      case 'get':
        return await getChat(accessToken, params);
      case 'update':
        return await updateChat(accessToken, params);
      case 'delete':
        return await deleteChat(accessToken, params);
      case 'list_messages':
        return await listChatMessages(accessToken, params);
      case 'get_message':
        return await getChatMessage(accessToken, params);
      case 'send_message':
        return await sendChatMessage(accessToken, params);
      case 'reply_to_message':
        return await replyToChatMessage(accessToken, params);
      case 'update_message':
        return await updateChatMessage(accessToken, params);
      case 'delete_message':
        return await deleteChatMessage(accessToken, params);
      case 'list_members':
        return await listChatMembers(accessToken, params);
      case 'add_member':
        return await addChatMember(accessToken, params);
      case 'remove_member':
        return await removeChatMember(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, create, get, update, delete, list_messages, get_message, send_message, reply_to_message, update_message, delete_message, list_members, add_member, remove_member` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in teams_chat ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in teams_chat operation: ${error.message}` }]
    };
  }
}

const MIN_ID_PREFIX = 24;

function commonPrefixLength(a, b) {
  const max = Math.min(a.length, b.length);
  let i = 0;
  while (i < max && a[i] === b[i]) i += 1;
  return i;
}

/**
 * Makes params.chatId usable: keeps a valid ID, repairs a garbled/truncated one
 * by its leading characters, or finds the chat from a name/email in
 * params.recipient. Mutates params. Returns { error } only when it is ambiguous.
 */
async function resolveChatParam(accessToken, params) {
  const recipient = params.recipient || params.to || params.userName;

  if (!params.chatId) {
    if (recipient) {
      const match = await findChatByName(accessToken, String(recipient));
      if (match.error) return { error: match.error };
      params.chatId = match.chatId;
    }
    return {};
  }

  try {
    const response = await callGraphAPI(accessToken, 'GET', 'me/chats', null, {
      $top: 50,
      $select: 'id'
    });
    const ids = (response.value || []).map((c) => c.id);
    if (ids.includes(params.chatId)) return {};

    const scored = ids
      .map((id) => ({ id, score: commonPrefixLength(id, params.chatId) }))
      .filter((c) => c.score >= MIN_ID_PREFIX)
      .sort((a, b) => b.score - a.score);
    if (scored.length > 0 && (scored.length === 1 || scored[0].score > scored[1].score)) {
      console.error(`Repaired chatId ${params.chatId} -> ${scored[0].id}`);
      params.chatId = scored[0].id;
    }
  } catch (error) {
    console.error('Could not verify chatId:', error.message);
  }
  return {};
}

/**
 * Newest messages across the user's most recently active chats.
 * Used when list_messages is called without a chatId.
 */
async function listLatestMessagesAcrossChats(accessToken, maxResults = 10) {
  const limit = Math.max(1, Math.min(Number(maxResults) || 10, 25));
  const allChats = await callGraphAPI(accessToken, 'GET', 'me/chats', null, {
    $top: 50,
    $expand: 'members',
    $select: 'id,topic,chatType,lastUpdatedDateTime'
  });
  const recentChats = (allChats.value || [])
    .sort((a, b) => new Date(b.lastUpdatedDateTime) - new Date(a.lastUpdatedDateTime))
    .slice(0, 8);

  const collected = [];
  await Promise.all(recentChats.map(async (chat) => {
    const name = chat.topic ||
      (chat.members || []).map((m) => m.displayName).filter(Boolean).join(', ') || 'Chat';
    try {
      const response = await callGraphAPI(accessToken, 'GET', `chats/${chat.id}/messages`, null, {
        $top: 5,
        $orderby: 'createdDateTime desc'
      });
      for (const message of response.value || []) {
        if (message.messageType && message.messageType !== 'message') continue;
        collected.push({ chat, name, message });
      }
    } catch (error) {
      console.error(`Could not read messages for chat ${chat.id}:`, error.message);
    }
  }));

  if (collected.length === 0) {
    return { content: [{ type: "text", text: "No recent Teams chat messages found." }] };
  }

  collected.sort((a, b) => new Date(b.message.createdDateTime) - new Date(a.message.createdDateTime));
  const lines = collected.slice(0, limit).map(({ chat, name, message }, index) => {
    const sender = message.from?.user?.displayName || message.from?.application?.displayName || 'Unknown';
    const preview = message.body?.content ? truncateContent(message.body.content, 200) : 'No content';
    return `${index + 1}. ${name} | ${sender} (${formatDate(message.createdDateTime)})\n   ${preview}\n   (chatId: ${chat.id})`;
  });

  return {
    content: [{
      type: "text",
      text: `Latest ${lines.length} Teams chat messages:\n\n${lines.join('\n\n')}`
    }]
  };
}

/**
 * List all chats the user is part of
 */
async function listChats(accessToken, params) {
  const { maxResults = 50 } = params;
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/chats',
    null,
    {
      $top: maxResults,
      $expand: 'members',
      $select: 'id,topic,webUrl,chatType,createdDateTime,lastUpdatedDateTime'
    }
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No chats found." }]
    };
  }
  
  // Format the chat list
  const chatList = response.value.map((chat, index) => {
    const chatType = getChatTypeName(chat.chatType);
    const memberCount = chat.members?.length || 0;
    
    // Create a descriptive name for the chat
    let chatName = chat.topic || 'Unnamed chat';
    if (!chat.topic) {
      if (chat.members && chat.members.length > 0) {
        // Create a name from first 3 members for untitled chats
        const memberNames = chat.members
          .filter(m => m.displayName)
          .map(m => m.displayName)
          .slice(0, 3);
        
        if (memberNames.length > 0) {
          chatName = memberNames.join(', ');
          if (chat.members.length > 3) {
            chatName += `, +${chat.members.length - 3} more`;
          }
        }
      }
    }
    
    return `${index + 1}. ${chatName}\n   ID: ${chat.id}\n   Type: ${chatType}\n   Members: ${memberCount}\n   Last updated: ${formatDate(chat.lastUpdatedDateTime)}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} chats:\n\n${chatList}` 
    }]
  };
}

/**
 * Create a new chat
 */
async function createChat(accessToken, params) {
  const { topic, members } = params;
  
  if (!members || !Array.isArray(members) || members.length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: members. Please provide an array of member identifiers (user IDs or emails)." 
      }]
    };
  }
  
  // Create the chat object
  const newChat = {
    chatType: members.length > 1 ? 'group' : 'oneOnOne',
    members: members.map(member => ({
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${member}')`
    }))
  };
  
  // Add the topic for group chats
  if (topic && members.length > 1) {
    newChat.topic = topic;
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    'chats',
    newChat
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Chat created successfully!\nID: ${response.id}\nType: ${getChatTypeName(response.chatType)}\nTopic: ${response.topic || 'Not specified'}` 
    }]
  };
}

/**
 * Get details about a specific chat
 */
async function getChat(accessToken, params) {
  const { chatId } = params;
  
  if (!chatId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: chatId" 
      }]
    };
  }
  
  const chat = await callGraphAPI(
    accessToken,
    'GET',
    `chats/${chatId}`,
    null,
    {
      $expand: 'members'
    }
  );
  
  // Format chat details
  const chatType = getChatTypeName(chat.chatType);
  const memberCount = chat.members?.length || 0;
  
  let chatDetails = `Chat: ${chat.topic || 'Untitled chat'}\n`;
  chatDetails += `ID: ${chat.id}\n`;
  chatDetails += `Type: ${chatType}\n`;
  chatDetails += `Created: ${formatDate(chat.createdDateTime)}\n`;
  chatDetails += `Last updated: ${formatDate(chat.lastUpdatedDateTime)}\n\n`;
  
  // Add member details
  if (chat.members && chat.members.length > 0) {
    chatDetails += `Members (${memberCount}):\n`;
    chat.members.forEach((member, index) => {
      const displayName = member.displayName || 'Unknown user';
      const email = member.email || 'No email';
      const roles = member.roles && member.roles.length > 0 ? member.roles.join(', ') : 'member';
      
      chatDetails += `${index + 1}. ${displayName}\n   Email: ${email}\n   Roles: ${roles}\n`;
    });
  }
  
  return {
    content: [{ type: "text", text: chatDetails }]
  };
}

/**
 * Update a chat (currently only topic can be updated)
 */
async function updateChat(accessToken, params) {
  const { chatId, topic } = params;
  
  if (!chatId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: chatId" 
      }]
    };
  }
  
  if (!topic) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing parameter: topic. Currently only the topic can be updated." 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'PATCH',
    `chats/${chatId}`,
    {
      topic
    }
  );
  
  return {
    content: [{ type: "text", text: "Chat updated successfully!" }]
  };
}

/**
 * Delete a chat
 */
async function deleteChat(accessToken, params) {
  const { chatId } = params;
  
  if (!chatId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: chatId" 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'DELETE',
    `chats/${chatId}`
  );
  
  return {
    content: [{ type: "text", text: "Chat deleted successfully!" }]
  };
}

/**
 * List messages in a chat
 */
async function listChatMessages(accessToken, params) {
  const { chatId, maxResults = 25 } = params;
  
  if (!chatId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: chatId" 
      }]
    };
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `chats/${chatId}/messages`,
    null,
    {
      $top: maxResults,
      $orderby: 'createdDateTime desc'
    }
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No messages found in this chat." }]
    };
  }
  
  // Format the message list
  const messageList = response.value.map((message, index) => {
    const sender = message.from?.user?.displayName || message.from?.user?.id || message.from?.application?.displayName || 'Unknown';
    const createdTime = formatDate(message.createdDateTime);
    const hasAttachments = message.attachments && message.attachments.length > 0;
    const contentPreview = message.body?.content ? truncateContent(message.body.content, 100) : 'No content';
    
    return `${index + 1}. From: ${sender} (${createdTime})\n   ID: ${message.id}\n   ${hasAttachments ? 'Has attachments' : 'No attachments'}\n   ${contentPreview}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} messages in chat (newest first):\n\n${messageList}` 
    }]
  };
}

/**
 * Get a specific chat message
 */
async function getChatMessage(accessToken, params) {
  const { chatId, messageId } = params;
  
  if (!chatId || !messageId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide chatId and messageId." 
      }]
    };
  }
  
  const message = await callGraphAPI(
    accessToken,
    'GET',
    `chats/${chatId}/messages/${messageId}`
  );
  
  const sender = message.from?.user?.displayName || message.from?.user?.id || message.from?.application?.displayName || 'Unknown';
  const createdTime = formatDate(message.createdDateTime);
  
  let content = 'No content';
  if (message.body?.content) {
    // Attempt to remove HTML from the content
    content = message.body.content
      .replace(/<[^>]*>/g, ' ') // Replace HTML tags with space
      .replace(/&nbsp;/g, ' ')  // Replace &nbsp; with space
      .replace(/\s+/g, ' ')     // Collapse multiple spaces
      .trim();
  }
  
  // Check for attachments
  let attachments = 'No attachments';
  if (message.attachments && message.attachments.length > 0) {
    attachments = message.attachments.map((attachment, index) => {
      return `   ${index + 1}. ${attachment.name || 'Unnamed'} (${attachment.contentType || 'Unknown type'})`;
    }).join('\n');
  }
  
  // Check if this is a reply to another message
  let replyInfo = '';
  if (message.replyToId) {
    replyInfo = `\nReply to message: ${message.replyToId}\n`;
  }
  
  const messageDetails = `From: ${sender}\nTime: ${createdTime}\nID: ${message.id}${replyInfo}\n\nContent:\n${content}\n\nAttachments:\n${attachments}`;
  
  return {
    content: [{ type: "text", text: messageDetails }]
  };
}

/**
 * Find an existing chat by person name, email or group topic.
 * Prefers a 1:1 chat; reports ambiguity or no match via `error`.
 */
async function findChatByName(accessToken, query) {
  // Display names often hold stray spaces ("Naveen  Vashist"), so compare on
  // whitespace-collapsed words, in any order.
  const normalize = (text) =>
    String(text || '').toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9@.\s-]/g, ' ').replace(/\s+/g, ' ').trim();
  const needle = normalize(query);
  const needleWords = needle.split(' ').filter(Boolean);
  const containsAllWords = (text) => {
    const haystack = normalize(text);
    return needleWords.length > 0 && needleWords.every((word) => haystack.includes(word));
  };

  const response = await callGraphAPI(accessToken, 'GET', 'me/chats', null, {
    $top: 50,
    $expand: 'members',
    $select: 'id,topic,chatType,lastUpdatedDateTime'
  });

  const label = (chat) =>
    chat.topic || (chat.members || []).map((m) => m.displayName).filter(Boolean).join(', ') || chat.id;

  const matches = (response.value || []).filter((chat) => {
    if (chat.topic && containsAllWords(chat.topic)) return true;
    return (chat.members || []).some(
      (m) => containsAllWords(m.displayName) || normalize(m.email) === needle
    );
  });

  const oneOnOne = matches.filter((c) => c.chatType === 'oneOnOne' && (c.members || []).length === 2);
  const candidates = oneOnOne.length > 0 ? oneOnOne : matches;

  if (candidates.length === 0) {
    return {
      error: `No existing chat found for "${query}". Use teams_chat create with the person's email in members to start one, then send_message with the new chatId.`
    };
  }
  if (candidates.length > 1) {
    const list = candidates.map((c) => `- ${label(c)} (chatId: ${c.id})`).join('\n');
    return { error: `Several chats match "${query}". Retry send_message with one of these chatIds:\n${list}` };
  }
  return { chatId: candidates[0].id };
}

/**
 * Send a message to a chat
 */
async function sendChatMessage(accessToken, params) {
  const { content, replyToId, attachments } = params;
  let { chatId } = params;

  // Allow addressing by person/group name or email instead of a chat ID.
  const recipient = params.recipient || params.to || params.userName || params.email;
  if (!chatId && recipient && content) {
    const match = await findChatByName(accessToken, String(recipient));
    if (match.error) {
      return { content: [{ type: "text", text: match.error }] };
    }
    chatId = match.chatId;
  }

  if (!chatId || !content) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide chatId and content." 
      }]
    };
  }
  
  const messageData = {
    body: {
      content,
      contentType: content.includes('<') ? 'html' : 'text'
    }
  };
  
  // Add reply information if this is a reply
  if (replyToId) {
    messageData.replyToId = replyToId;
  }
  
  // Add attachments if present
  if (attachments && Array.isArray(attachments) && attachments.length > 0) {
    messageData.attachments = attachments.map(attachment => {
      const { name, contentUrl, contentType } = attachment;
      return { name, contentUrl, contentType };
    });
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    `chats/${chatId}/messages`,
    messageData
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Message sent successfully!\nID: ${response.id}` 
    }]
  };
}

/**
 * Reply to a specific message in a group or 1:1 chat (quoted reply in Teams UI).
 * Uses Graph replyWithQuote — replyToId on send_message does not thread in chats.
 */
async function replyToChatMessage(accessToken, params) {
  const { chatId, messageId, messageIds, content } = params;

  const ids = Array.isArray(messageIds) && messageIds.length > 0
    ? messageIds
    : messageId
      ? [messageId]
      : [];

  if (!chatId || ids.length === 0 || !content) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameters. Please provide chatId, messageId (or messageIds), and content.",
      }],
    };
  }

  if (ids.length > 10) {
    return {
      content: [{
        type: "text",
        text: "Too many messageIds. replyWithQuote supports at most 10 quoted messages.",
      }],
    };
  }

  const replyData = {
    messageIds: ids,
    replyMessage: {
      body: {
        content,
        contentType: content.includes('<') ? 'html' : 'text',
      },
    },
  };

  const response = await callGraphAPI(
    accessToken,
    'POST',
    `chats/${chatId}/messages/replyWithQuote`,
    replyData,
  );

  const quoted = ids.length === 1 ? ids[0] : ids.join(', ');
  return {
    content: [{
      type: "text",
      text: `Reply sent successfully (quoted message${ids.length > 1 ? 's' : ''}: ${quoted})!\nID: ${response.id}`,
    }],
  };
}

/**
 * Update a chat message
 */
async function updateChatMessage(accessToken, params) {
  const { chatId, messageId, content } = params;
  
  if (!chatId || !messageId || !content) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide chatId, messageId, and content." 
      }]
    };
  }
  
  const messageData = {
    body: {
      content,
      contentType: content.includes('<') ? 'html' : 'text'
    }
  };
  
  await callGraphAPI(
    accessToken,
    'PATCH',
    `chats/${chatId}/messages/${messageId}`,
    messageData
  );
  
  return {
    content: [{ type: "text", text: "Message updated successfully!" }]
  };
}

/**
 * Delete a chat message
 */
async function deleteChatMessage(accessToken, params) {
  const { chatId, messageId } = params;
  
  if (!chatId || !messageId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide chatId and messageId." 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'DELETE',
    `chats/${chatId}/messages/${messageId}`
  );
  
  return {
    content: [{ type: "text", text: "Message deleted successfully!" }]
  };
}

/**
 * List members of a chat
 */
async function listChatMembers(accessToken, params) {
  const { chatId } = params;
  
  if (!chatId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: chatId" 
      }]
    };
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `chats/${chatId}/members`
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No members found in this chat." }]
    };
  }
  
  // Format the member list
  const memberList = response.value.map((member, index) => {
    const displayName = member.displayName || 'Unknown user';
    const email = member.email || 'No email';
    const userId = member.userId || 'No ID';
    const roles = member.roles && member.roles.length > 0 ? member.roles.join(', ') : 'member';
    
    return `${index + 1}. ${displayName}\n   Email: ${email}\n   ID: ${userId}\n   Roles: ${roles}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} members in chat:\n\n${memberList}` 
    }]
  };
}

/**
 * Add a member to a chat
 */
async function addChatMember(accessToken, params) {
  const { chatId, userId, email, roles = ['member'] } = params;
  
  if (!chatId || (!userId && !email)) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide chatId and either userId or email." 
      }]
    };
  }
  
  // Determine the member to add
  let userIdentifier;
  if (userId) {
    userIdentifier = userId;
  } else if (email) {
    // Try to resolve email to user ID
    try {
      const userResponse = await callGraphAPI(
        accessToken,
        'GET',
        'users',
        null,
        {
          $filter: `mail eq '${email}' or userPrincipalName eq '${email}'`,
          $select: 'id'
        }
      );
      
      if (userResponse.value && userResponse.value.length > 0) {
        userIdentifier = userResponse.value[0].id;
      } else {
        return {
          content: [{ type: "text", text: `Unable to find a user with email: ${email}` }]
        };
      }
    } catch (error) {
      console.error('Error resolving email to user ID:', error);
      return {
        content: [{ type: "text", text: `Error resolving email to user ID: ${error.message}` }]
      };
    }
  }
  
  const memberData = {
    '@odata.type': '#microsoft.graph.aadUserConversationMember',
    roles: roles,
    'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userIdentifier}')`
  };
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    `chats/${chatId}/members`,
    memberData
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Member added successfully!\nID: ${response.id}` 
    }]
  };
}

/**
 * Remove a member from a chat
 */
async function removeChatMember(accessToken, params) {
  const { chatId, memberId } = params;
  
  if (!chatId || !memberId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide chatId and memberId." 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'DELETE',
    `chats/${chatId}/members/${memberId}`
  );
  
  return {
    content: [{ type: "text", text: "Member removed successfully!" }]
  };
}

/**
 * Helper function to get a friendly name for chat types
 */
function getChatTypeName(chatType) {
  switch (chatType) {
    case 'oneOnOne':
      return 'One-on-One Chat';
    case 'group':
      return 'Group Chat';
    case 'meeting':
      return 'Meeting Chat';
    default:
      return chatType || 'Unknown';
  }
}

/**
 * Helper function to format dates
 */
function formatDate(dateString) {
  if (!dateString) return 'Unknown';
  try {
    return new Date(dateString).toLocaleString();
  } catch (_error) {
    return dateString;
  }
}

/**
 * Helper function to truncate content to a specific length
 */
function truncateContent(content, maxLength) {
  if (!content) return 'No content';
  
  // Remove HTML tags
  const plainText = content
    .replace(/<[^>]*>/g, ' ')  // Replace HTML tags with space
    .replace(/&nbsp;/g, ' ')   // Replace &nbsp; with space
    .replace(/\s+/g, ' ')      // Collapse multiple spaces
    .trim();
  
  if (plainText.length <= maxLength) {
    return plainText;
  }
  
  return plainText.substring(0, maxLength) + '...';
}

// Export the handler
module.exports = handleTeamsChat;