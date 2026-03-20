/**
 * Domain types for the enterprise meeting search application.
 */

/** Represents a single chat message search result. */
export interface ChatMessageResult {
  /** Unique identifier of the message */
  messageId: string;
  /** Display name of the message sender */
  senderName: string;
  /** User principal name (email) of the sender */
  senderEmail: string | null;
  /** Plain-text body of the message */
  body: string;
  /** ISO 8601 timestamp when the message was created */
  createdDateTime: string;
  /** ISO 8601 timestamp of the last modification, if any */
  lastModifiedDateTime: string | null;
  /** Context describing where the message lives */
  context: ChatContext | ChannelContext;
  /** Deep-link URL to the message in Teams */
  webUrl: string | null;
  /** ID of the parent Teams meeting, when the chat is meeting-associated */
  meetingId: string | null;
}

/** Context for a message that lives in a chat (1:1, group, or meeting chat). */
export interface ChatContext {
  type: 'chat';
  /** ID of the chat thread */
  chatId: string;
  /** Human-readable topic/name of the chat, if available */
  chatTopic: string | null;
  /** Category of the chat */
  chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknownFutureValue';
}

/** Context for a message that lives in a Teams channel. */
export interface ChannelContext {
  type: 'channel';
  /** ID of the team */
  teamId: string;
  /** Display name of the team */
  teamName: string | null;
  /** ID of the channel */
  channelId: string;
  /** Display name of the channel */
  channelName: string | null;
}

/** Options that control how keywords are combined during a search. */
export interface KeywordSearchOptions {
  /**
   * List of keywords to search for.
   * When `operator` is `'AND'` every keyword must appear in the message.
   * When `operator` is `'OR'`  any keyword may appear.
   */
  keywords: string[];
  /** Logical operator used to combine multiple keywords. Defaults to `'AND'`. */
  operator?: 'AND' | 'OR';
}

/** Raw Graph API chat object (minimal fields used by this module). */
export interface GraphChat {
  id: string;
  topic: string | null;
  chatType: string;
  onlineMeetingInfo?: {
    organizerOrganizationId?: string | null;
    joinWebUrl?: string | null;
    calendarEventId?: string | null;
  } | null;
}

/** Raw Graph API channel object (minimal fields used by this module). */
export interface GraphChannel {
  id: string;
  displayName: string | null;
}

/** Raw Graph API team object (minimal fields used by this module). */
export interface GraphTeam {
  id: string;
  displayName: string | null;
}

/** Raw Graph API chat message object (minimal fields used by this module). */
export interface GraphChatMessage {
  id: string;
  body?: {
    content: string | null;
    contentType?: string;
  };
  from?: {
    user?: {
      displayName: string | null;
      userIdentityType?: string;
      id?: string | null;
    } | null;
  } | null;
  createdDateTime: string;
  lastModifiedDateTime: string | null;
  webUrl: string | null;
  chatId?: string | null;
  channelIdentity?: {
    teamId?: string;
    channelId?: string;
  } | null;
}

/** Paginated Graph API response wrapper. */
export interface GraphListResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
}
