import { Client } from '@microsoft/microsoft-graph-client';
import {
  ChatMessageResult,
  ChatContext,
  ChannelContext,
  GraphChat,
  GraphChannel,
  GraphTeam,
  GraphChatMessage,
  GraphListResponse,
  KeywordSearchOptions,
} from '../types';

/**
 * Service for searching Teams chat messages across meeting chats, 1:1 / group
 * chats, and channel conversations using Microsoft Graph API.
 */
export class TeamsChatSearchService {
  private readonly client: Client;

  constructor(client: Client) {
    this.client = client;
  }

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------

  /**
   * Search for messages across all chats (meeting, 1:1, and group) that the
   * service account has access to.
   *
   * @param options - Keywords and logical operator for the search.
   * @returns Array of matching {@link ChatMessageResult} objects.
   */
  async searchChatMessages(
    options: KeywordSearchOptions,
  ): Promise<ChatMessageResult[]> {
    const operator = options.operator ?? 'AND';
    const chats = await this.listAllChats();
    const results: ChatMessageResult[] = [];

    await Promise.all(
      chats.map(async (chat) => {
        const messages = await this.listChatMessages(chat.id);
        for (const message of messages) {
          if (this.messageMatchesKeywords(message, options.keywords, operator)) {
            results.push(this.mapChatMessage(message, chat));
          }
        }
      }),
    );

    return results;
  }

  /**
   * Search for messages in Teams channel conversations across all teams and
   * channels the service account has access to.
   *
   * @param options - Keywords and logical operator for the search.
   * @returns Array of matching {@link ChatMessageResult} objects.
   */
  async searchChannelMessages(
    options: KeywordSearchOptions,
  ): Promise<ChatMessageResult[]> {
    const operator = options.operator ?? 'AND';
    const teams = await this.listAllTeams();
    const results: ChatMessageResult[] = [];

    await Promise.all(
      teams.map(async (team) => {
        const channels = await this.listTeamChannels(team.id);
        await Promise.all(
          channels.map(async (channel) => {
            const messages = await this.listChannelMessages(
              team.id,
              channel.id,
            );
            for (const message of messages) {
              if (
                this.messageMatchesKeywords(message, options.keywords, operator)
              ) {
                results.push(this.mapChannelMessage(message, team, channel));
              }
            }
          }),
        );
      }),
    );

    return results;
  }

  /**
   * Convenience method that searches both chat messages and channel messages
   * and returns the combined results.
   *
   * @param options - Keywords and logical operator for the search.
   * @returns Combined array of matching {@link ChatMessageResult} objects.
   */
  async searchAllMessages(
    options: KeywordSearchOptions,
  ): Promise<ChatMessageResult[]> {
    const [chatResults, channelResults] = await Promise.all([
      this.searchChatMessages(options),
      this.searchChannelMessages(options),
    ]);
    return [...chatResults, ...channelResults];
  }

  // ---------------------------------------------------------------------------
  // Graph API helpers – chats
  // ---------------------------------------------------------------------------

  /** Retrieve all chats accessible to the service account (paginated). */
  async listAllChats(): Promise<GraphChat[]> {
    return this.fetchAllPages<GraphChat>('/chats?$expand=onlineMeetingInfo');
  }

  /** Retrieve all messages in a given chat thread (paginated). */
  async listChatMessages(chatId: string): Promise<GraphChatMessage[]> {
    return this.fetchAllPages<GraphChatMessage>(
      `/chats/${chatId}/messages`,
    );
  }

  // ---------------------------------------------------------------------------
  // Graph API helpers – teams / channels
  // ---------------------------------------------------------------------------

  /** Retrieve all teams accessible to the service account (paginated). */
  async listAllTeams(): Promise<GraphTeam[]> {
    return this.fetchAllPages<GraphTeam>('/teams');
  }

  /** Retrieve all channels in a given team (paginated). */
  async listTeamChannels(teamId: string): Promise<GraphChannel[]> {
    return this.fetchAllPages<GraphChannel>(`/teams/${teamId}/channels`);
  }

  /** Retrieve all messages in a given channel (paginated). */
  async listChannelMessages(
    teamId: string,
    channelId: string,
  ): Promise<GraphChatMessage[]> {
    return this.fetchAllPages<GraphChatMessage>(
      `/teams/${teamId}/channels/${channelId}/messages`,
    );
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  /**
   * Fetches all pages of a paginated Graph API list endpoint.
   *
   * @param path - Relative path (e.g. `/chats`).
   * @returns All items from all pages combined.
   */
  private async fetchAllPages<T>(path: string): Promise<T[]> {
    const items: T[] = [];
    let nextLink: string | undefined = path;

    while (nextLink) {
      const response: GraphListResponse<T> =
        await this.client.api(nextLink).get();
      items.push(...response.value);
      nextLink = response['@odata.nextLink'];
    }

    return items;
  }

  /**
   * Returns `true` when the message body satisfies the keyword filter.
   *
   * - `AND`: every keyword must appear (case-insensitive) in the body.
   * - `OR`:  at least one keyword must appear.
   */
  private messageMatchesKeywords(
    message: GraphChatMessage,
    keywords: string[],
    operator: 'AND' | 'OR',
  ): boolean {
    if (keywords.length === 0) return true;

    const body = (message.body?.content ?? '').toLowerCase();

    if (operator === 'AND') {
      return keywords.every((kw) => body.includes(kw.toLowerCase()));
    }
    return keywords.some((kw) => body.includes(kw.toLowerCase()));
  }

  /**
   * Maps a raw Graph chat message to a {@link ChatMessageResult} enriched with
   * chat context and meeting correlation.
   */
  private mapChatMessage(
    message: GraphChatMessage,
    chat: GraphChat,
  ): ChatMessageResult {
    const context: ChatContext = {
      type: 'chat',
      chatId: chat.id,
      chatTopic: chat.topic,
      chatType: chat.chatType as ChatContext['chatType'],
    };

    // Correlate meeting chats to their parent meeting.
    // The `calendarEventId` inside `onlineMeetingInfo` is the most reliable
    // identifier available via Graph for meeting-associated chats.
    const meetingId =
      chat.onlineMeetingInfo?.calendarEventId ??
      chat.onlineMeetingInfo?.joinWebUrl?.match(/meetup-join\/([^/]+)/)?.[1] ??
      null;

    return {
      messageId: message.id,
      senderName: message.from?.user?.displayName ?? 'Unknown',
      senderEmail: message.from?.user?.id ?? null,
      body: message.body?.content ?? '',
      createdDateTime: message.createdDateTime,
      lastModifiedDateTime: message.lastModifiedDateTime,
      context,
      webUrl: message.webUrl,
      meetingId: chat.chatType === 'meeting' ? meetingId : null,
    };
  }

  /**
   * Maps a raw Graph channel message to a {@link ChatMessageResult} enriched
   * with channel/team context.
   */
  private mapChannelMessage(
    message: GraphChatMessage,
    team: GraphTeam,
    channel: GraphChannel,
  ): ChatMessageResult {
    const context: ChannelContext = {
      type: 'channel',
      teamId: team.id,
      teamName: team.displayName,
      channelId: channel.id,
      channelName: channel.displayName,
    };

    return {
      messageId: message.id,
      senderName: message.from?.user?.displayName ?? 'Unknown',
      senderEmail: message.from?.user?.id ?? null,
      body: message.body?.content ?? '',
      createdDateTime: message.createdDateTime,
      lastModifiedDateTime: message.lastModifiedDateTime,
      context,
      webUrl: message.webUrl,
      meetingId: null,
    };
  }
}
