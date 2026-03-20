import { Client } from '@microsoft/microsoft-graph-client';
import { TeamsChatSearchService } from '../../src/search/teamsChatSearch';
import {
  GraphChat,
  GraphChannel,
  GraphTeam,
  GraphChatMessage,
  ChatContext,
  ChannelContext,
} from '../../src/types';

// ---------------------------------------------------------------------------
// Test fixtures
// ---------------------------------------------------------------------------

const MEETING_CHAT: GraphChat = {
  id: 'chat-meeting-1',
  topic: 'Project Kickoff',
  chatType: 'meeting',
  onlineMeetingInfo: {
    calendarEventId: 'event-abc123',
    joinWebUrl:
      'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc/0?context={}',
  },
};

const ONE_ON_ONE_CHAT: GraphChat = {
  id: 'chat-1on1-1',
  topic: null,
  chatType: 'oneOnOne',
  onlineMeetingInfo: null,
};

const GROUP_CHAT: GraphChat = {
  id: 'chat-group-1',
  topic: 'WSIB Updates',
  chatType: 'group',
  onlineMeetingInfo: null,
};

const TEAM: GraphTeam = {
  id: 'team-1',
  displayName: 'Engineering',
};

const CHANNEL: GraphChannel = {
  id: 'channel-1',
  displayName: 'General',
};

const makeMessage = (
  id: string,
  content: string,
  chatId?: string,
): GraphChatMessage => ({
  id,
  body: { content, contentType: 'text' },
  from: {
    user: { displayName: 'Alice Smith', id: 'alice@contoso.com' },
  },
  createdDateTime: '2024-06-01T10:00:00Z',
  lastModifiedDateTime: null,
  webUrl: `https://teams.microsoft.com/l/message/${id}`,
  chatId: chatId ?? null,
});

// Meeting chat messages
const MEETING_MSG_MATCH = makeMessage(
  'msg-m1',
  'The PTU allocation for WSIB has been reviewed.',
  MEETING_CHAT.id,
);
const MEETING_MSG_NO_MATCH = makeMessage(
  'msg-m2',
  'See you at the standup tomorrow.',
  MEETING_CHAT.id,
);

// 1:1 chat messages
const ONE_ON_ONE_MSG_MATCH = makeMessage(
  'msg-1on1-1',
  'Did you check the PTU numbers?',
  ONE_ON_ONE_CHAT.id,
);
const ONE_ON_ONE_MSG_NO_MATCH = makeMessage(
  'msg-1on1-2',
  'Lunch at noon?',
  ONE_ON_ONE_CHAT.id,
);

// Group chat messages
const GROUP_MSG_MATCH = makeMessage(
  'msg-g1',
  'WSIB PTU report is ready for review.',
  GROUP_CHAT.id,
);

// Channel messages
const CHANNEL_MSG_MATCH = makeMessage('msg-ch1', 'PTU review complete.');
const CHANNEL_MSG_NO_MATCH = makeMessage('msg-ch2', 'Standup notes attached.');

// ---------------------------------------------------------------------------
// Mock Graph Client factory
// ---------------------------------------------------------------------------

/**
 * Builds a minimal mock of the Graph `Client` that routes API calls to the
 * provided handler map keyed by path prefix.
 * Keys are matched longest-first so that specific paths take precedence over
 * shorter prefixes (e.g. `/chats/id/messages` before `/chats`).
 */
function buildMockClient(
  handlers: Record<string, unknown>,
): Client {
  const sortedKeys = Object.keys(handlers).sort((a, b) => b.length - a.length);

  const mockGet = jest.fn((path: string) => {
    for (const prefix of sortedKeys) {
      if (path.startsWith(prefix)) {
        const data = handlers[prefix];
        return Promise.resolve(
          Array.isArray(data) ? { value: data } : data,
        );
      }
    }
    return Promise.resolve({ value: [] });
  });

  return {
    api: (path: string) => ({
      get: () => mockGet(path),
    }),
  } as unknown as Client;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('TeamsChatSearchService', () => {
  // -------------------------------------------------------------------------
  // searchChatMessages
  // -------------------------------------------------------------------------

  describe('searchChatMessages', () => {
    it('returns meeting-chat messages that contain all AND keywords', async () => {
      const client = buildMockClient({
        '/chats': [MEETING_CHAT, ONE_ON_ONE_CHAT, GROUP_CHAT],
        [`/chats/${MEETING_CHAT.id}/messages`]: [
          MEETING_MSG_MATCH,
          MEETING_MSG_NO_MATCH,
        ],
        [`/chats/${ONE_ON_ONE_CHAT.id}/messages`]: [ONE_ON_ONE_MSG_NO_MATCH],
        [`/chats/${GROUP_CHAT.id}/messages`]: [],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({
        keywords: ['PTU', 'WSIB'],
        operator: 'AND',
      });

      expect(results).toHaveLength(1);
      expect(results[0].messageId).toBe('msg-m1');
      expect(results[0].senderName).toBe('Alice Smith');
      expect(results[0].body).toContain('PTU');
      expect(results[0].body).toContain('WSIB');
    });

    it('correlates a meeting chat message to its parent meeting ID', async () => {
      const client = buildMockClient({
        '/chats': [MEETING_CHAT],
        [`/chats/${MEETING_CHAT.id}/messages`]: [MEETING_MSG_MATCH],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({ keywords: ['PTU'] });

      expect(results).toHaveLength(1);
      expect(results[0].meetingId).toBe('event-abc123');
      const ctx = results[0].context as ChatContext;
      expect(ctx.type).toBe('chat');
      expect(ctx.chatType).toBe('meeting');
    });

    it('does not set meetingId for non-meeting chats', async () => {
      const client = buildMockClient({
        '/chats': [ONE_ON_ONE_CHAT],
        [`/chats/${ONE_ON_ONE_CHAT.id}/messages`]: [ONE_ON_ONE_MSG_MATCH],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({ keywords: ['PTU'] });

      expect(results[0].meetingId).toBeNull();
      const ctx = results[0].context as ChatContext;
      expect(ctx.chatType).toBe('oneOnOne');
    });

    it('returns 1:1 and group chat messages that match OR keywords', async () => {
      const client = buildMockClient({
        '/chats': [ONE_ON_ONE_CHAT, GROUP_CHAT],
        [`/chats/${ONE_ON_ONE_CHAT.id}/messages`]: [
          ONE_ON_ONE_MSG_MATCH,
          ONE_ON_ONE_MSG_NO_MATCH,
        ],
        [`/chats/${GROUP_CHAT.id}/messages`]: [GROUP_MSG_MATCH],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({
        keywords: ['PTU', 'WSIB'],
        operator: 'OR',
      });

      // ONE_ON_ONE_MSG_MATCH has PTU; GROUP_MSG_MATCH has both
      expect(results.length).toBeGreaterThanOrEqual(2);
      const ids = results.map((r) => r.messageId);
      expect(ids).toContain('msg-1on1-1');
      expect(ids).toContain('msg-g1');
    });

    it('returns an empty array when no messages match', async () => {
      const client = buildMockClient({
        '/chats': [ONE_ON_ONE_CHAT],
        [`/chats/${ONE_ON_ONE_CHAT.id}/messages`]: [ONE_ON_ONE_MSG_NO_MATCH],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({ keywords: ['PTU'] });

      expect(results).toHaveLength(0);
    });

    it('returns all messages when no keywords are supplied', async () => {
      const client = buildMockClient({
        '/chats': [MEETING_CHAT],
        [`/chats/${MEETING_CHAT.id}/messages`]: [
          MEETING_MSG_MATCH,
          MEETING_MSG_NO_MATCH,
        ],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({ keywords: [] });

      expect(results).toHaveLength(2);
    });

    it('includes chat topic and webUrl in the result', async () => {
      const client = buildMockClient({
        '/chats': [GROUP_CHAT],
        [`/chats/${GROUP_CHAT.id}/messages`]: [GROUP_MSG_MATCH],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({ keywords: ['WSIB'] });

      expect(results).toHaveLength(1);
      const ctx = results[0].context as ChatContext;
      expect(ctx.chatTopic).toBe('WSIB Updates');
      expect(results[0].webUrl).toContain('msg-g1');
      expect(results[0].createdDateTime).toBe('2024-06-01T10:00:00Z');
    });
  });

  // -------------------------------------------------------------------------
  // searchChannelMessages
  // -------------------------------------------------------------------------

  describe('searchChannelMessages', () => {
    it('returns channel messages matching AND keywords', async () => {
      const client = buildMockClient({
        '/teams': [TEAM],
        [`/teams/${TEAM.id}/channels`]: [CHANNEL],
        [`/teams/${TEAM.id}/channels/${CHANNEL.id}/messages`]: [
          CHANNEL_MSG_MATCH,
          CHANNEL_MSG_NO_MATCH,
        ],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChannelMessages({ keywords: ['PTU'] });

      expect(results).toHaveLength(1);
      expect(results[0].messageId).toBe('msg-ch1');
      const ctx = results[0].context as ChannelContext;
      expect(ctx.type).toBe('channel');
      expect(ctx.teamId).toBe(TEAM.id);
      expect(ctx.teamName).toBe('Engineering');
      expect(ctx.channelId).toBe(CHANNEL.id);
      expect(ctx.channelName).toBe('General');
    });

    it('sets meetingId to null for channel messages', async () => {
      const client = buildMockClient({
        '/teams': [TEAM],
        [`/teams/${TEAM.id}/channels`]: [CHANNEL],
        [`/teams/${TEAM.id}/channels/${CHANNEL.id}/messages`]: [
          CHANNEL_MSG_MATCH,
        ],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChannelMessages({ keywords: ['PTU'] });

      expect(results[0].meetingId).toBeNull();
    });

    it('returns empty array when no channels match keyword', async () => {
      const client = buildMockClient({
        '/teams': [TEAM],
        [`/teams/${TEAM.id}/channels`]: [CHANNEL],
        [`/teams/${TEAM.id}/channels/${CHANNEL.id}/messages`]: [
          CHANNEL_MSG_NO_MATCH,
        ],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChannelMessages({ keywords: ['PTU'] });

      expect(results).toHaveLength(0);
    });
  });

  // -------------------------------------------------------------------------
  // searchAllMessages
  // -------------------------------------------------------------------------

  describe('searchAllMessages', () => {
    it('combines chat and channel results', async () => {
      const client = buildMockClient({
        '/chats': [MEETING_CHAT],
        [`/chats/${MEETING_CHAT.id}/messages`]: [MEETING_MSG_MATCH],
        '/teams': [TEAM],
        [`/teams/${TEAM.id}/channels`]: [CHANNEL],
        [`/teams/${TEAM.id}/channels/${CHANNEL.id}/messages`]: [
          CHANNEL_MSG_MATCH,
        ],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchAllMessages({ keywords: ['PTU'] });

      expect(results.length).toBe(2);
      const types = results.map((r) => r.context.type);
      expect(types).toContain('chat');
      expect(types).toContain('channel');
    });
  });

  // -------------------------------------------------------------------------
  // Pagination
  // -------------------------------------------------------------------------

  describe('pagination', () => {
    it('follows @odata.nextLink to retrieve all pages', async () => {
      const page1 = {
        value: [MEETING_MSG_MATCH],
        '@odata.nextLink': `/chats/${MEETING_CHAT.id}/messages?$skiptoken=page2`,
      };
      const page2 = { value: [MEETING_MSG_NO_MATCH] };

      const mockGet = jest.fn()
        .mockResolvedValueOnce({ value: [MEETING_CHAT] }) // /chats
        .mockResolvedValueOnce(page1) // first page of messages
        .mockResolvedValueOnce(page2); // second page of messages

      const client = {
        api: (path: string): { get: () => Promise<unknown> } => ({ get: () => mockGet(path) }),
      } as unknown as Client;

      const svc = new TeamsChatSearchService(client);
      const results = await svc.searchChatMessages({ keywords: [] });

      expect(results).toHaveLength(2);
    });
  });

  // -------------------------------------------------------------------------
  // Meeting ID correlation via joinWebUrl fallback
  // -------------------------------------------------------------------------

  describe('meeting ID correlation', () => {
    it('falls back to joinWebUrl when calendarEventId is absent', async () => {
      const chatWithJoinUrl: GraphChat = {
        id: 'chat-meeting-2',
        topic: 'Fallback Meeting',
        chatType: 'meeting',
        onlineMeetingInfo: {
          calendarEventId: null,
          joinWebUrl:
            'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_XYZ/0',
        },
      };
      const msg = makeMessage('msg-x1', 'PTU discussion', chatWithJoinUrl.id);

      const client = buildMockClient({
        '/chats': [chatWithJoinUrl],
        [`/chats/${chatWithJoinUrl.id}/messages`]: [msg],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({ keywords: ['PTU'] });

      expect(results).toHaveLength(1);
      expect(results[0].meetingId).toBe('19%3Ameeting_XYZ');
    });

    it('returns null meetingId when onlineMeetingInfo is absent', async () => {
      const chatNoMeetingInfo: GraphChat = {
        id: 'chat-meeting-3',
        topic: 'No Info',
        chatType: 'meeting',
        onlineMeetingInfo: null,
      };
      const msg = makeMessage('msg-x2', 'PTU review', chatNoMeetingInfo.id);

      const client = buildMockClient({
        '/chats': [chatNoMeetingInfo],
        [`/chats/${chatNoMeetingInfo.id}/messages`]: [msg],
      });
      const svc = new TeamsChatSearchService(client);

      const results = await svc.searchChatMessages({ keywords: ['PTU'] });

      expect(results[0].meetingId).toBeNull();
    });
  });
});
