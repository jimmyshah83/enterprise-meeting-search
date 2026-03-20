import { EmailSearchService } from '../../src/email/emailSearchService';
import { GraphClientWrapper } from '../../src/graph/graphClient';
import { GraphMessage } from '../../src/email/emailSearchTypes';

/** Helper to build a minimal, valid Graph message. */
function makeGraphMessage(overrides: Partial<GraphMessage> = {}): GraphMessage {
  return {
    id: 'msg-default',
    subject: 'Weekly standup',
    bodyPreview: 'Agenda for this week.',
    receivedDateTime: '2024-01-15T10:00:00Z',
    from: { emailAddress: { address: 'alice@example.com', name: 'Alice' } },
    toRecipients: [
      { emailAddress: { address: 'bob@example.com', name: 'Bob' } },
    ],
    ccRecipients: [],
    hasAttachments: false,
    categories: [],
    ...overrides,
  };
}

/** Creates a mock GraphClientWrapper whose getMessages resolves to the supplied messages. */
function makeMockGraphClient(messages: GraphMessage[]): GraphClientWrapper {
  const mock = {
    getMessages: jest.fn().mockResolvedValue({ value: messages }),
  } as unknown as GraphClientWrapper;
  return mock;
}

describe('EmailSearchService', () => {
  describe('search – basic keyword search', () => {
    it('returns mapped results for a keyword search', async () => {
      const messages = [
        makeGraphMessage({ id: 'msg-1', subject: 'Budget review' }),
        makeGraphMessage({ id: 'msg-2', subject: 'Q2 planning invite' }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({ keywords: ['budget', 'planning'] });

      expect(response.results).toHaveLength(2);
      expect(response.totalCount).toBe(2);
      expect(response.results[0].id).toBe('msg-1');
      expect(response.results[1].id).toBe('msg-2');
    });

    it('passes the $search query parameter to the graph client', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await service.search({ keywords: ['standup', 'review'] });

      expect(client.getMessages).toHaveBeenCalledWith(
        expect.objectContaining({
          $search: expect.stringContaining('"standup"'),
        }),
      );
    });

    it('returns an empty result set when there are no messages', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      const response = await service.search({ keywords: ['nothing'] });

      expect(response.results).toHaveLength(0);
      expect(response.totalCount).toBe(0);
    });

    it('strips special characters from keywords before building $search', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await service.search({ keywords: ['hello"world', 'test; DROP TABLE--'] });

      const call = (client.getMessages as jest.Mock).mock.calls[0][0];
      // Double-quotes are stripped; the alphanumeric part is still wrapped in quotes.
      expect(call.$search).toContain('"helloworld"');
      // Semicolons and other non-safe chars are stripped; spaces and hyphens are kept.
      expect(call.$search).toContain('"test DROP TABLE--"');
    });
  });

  describe('search – date range filter', () => {
    it('includes receivedDateTime ge/le clauses in $filter', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      const dateFrom = new Date('2024-01-01T00:00:00Z');
      const dateTo = new Date('2024-01-31T23:59:59Z');

      await service.search({ dateFrom, dateTo });

      expect(client.getMessages).toHaveBeenCalledWith(
        expect.objectContaining({
          $filter: expect.stringContaining('receivedDateTime ge'),
        }),
      );
      expect(client.getMessages).toHaveBeenCalledWith(
        expect.objectContaining({
          $filter: expect.stringContaining('receivedDateTime le'),
        }),
      );
    });

    it('does not include $filter when no date or sender is supplied', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await service.search({ keywords: ['test'] });

      const call = (client.getMessages as jest.Mock).mock.calls[0][0];
      expect(call).not.toHaveProperty('$filter');
    });
  });

  describe('search – sender filter', () => {
    it('includes sender address in $filter', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await service.search({ sender: 'alice@example.com' });

      expect(client.getMessages).toHaveBeenCalledWith(
        expect.objectContaining({
          $filter: expect.stringContaining("from/emailAddress/address eq 'alice@example.com'"),
        }),
      );
    });

    it('throws for an invalid sender email address', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await expect(service.search({ sender: "invalid'; DROP TABLE--" })).rejects.toThrow(
        'Invalid email address supplied for filter',
      );
    });
  });

  describe('search – recipient filter (post-filter)', () => {
    it('keeps only messages where the target recipient appears in to/cc', async () => {
      const messages = [
        makeGraphMessage({
          id: 'msg-1',
          toRecipients: [{ emailAddress: { address: 'carol@example.com' } }],
        }),
        makeGraphMessage({
          id: 'msg-2',
          toRecipients: [{ emailAddress: { address: 'dave@example.com' } }],
        }),
        makeGraphMessage({
          id: 'msg-3',
          ccRecipients: [{ emailAddress: { address: 'carol@example.com' } }],
        }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({ recipients: ['carol@example.com'] });

      expect(response.results.map((r) => r.id)).toEqual(['msg-1', 'msg-3']);
    });

    it('is case-insensitive when matching recipients', async () => {
      const messages = [
        makeGraphMessage({
          id: 'msg-1',
          toRecipients: [{ emailAddress: { address: 'Carol@Example.COM' } }],
        }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({ recipients: ['carol@example.com'] });

      expect(response.results).toHaveLength(1);
    });
  });

  describe('search – maxResults', () => {
    it('passes maxResults as $top', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await service.search({ maxResults: 10 });

      expect(client.getMessages).toHaveBeenCalledWith(
        expect.objectContaining({ $top: 10 }),
      );
    });

    it('defaults to 25 when maxResults is not provided', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await service.search({});

      expect(client.getMessages).toHaveBeenCalledWith(
        expect.objectContaining({ $top: 25 }),
      );
    });
  });

  describe('search – attachment expansion', () => {
    it('includes $expand=attachments when includeAttachments is true', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await service.search({ includeAttachments: true });

      expect(client.getMessages).toHaveBeenCalledWith(
        expect.objectContaining({ $expand: 'attachments' }),
      );
    });

    it('does not include $expand when includeAttachments is false', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);

      await service.search({ includeAttachments: false });

      const call = (client.getMessages as jest.Mock).mock.calls[0][0];
      expect(call).not.toHaveProperty('$expand');
    });

    it('maps attachment metadata correctly', async () => {
      const messages = [
        makeGraphMessage({
          id: 'msg-1',
          hasAttachments: true,
          attachments: [
            { id: 'att-1', name: 'agenda.pdf', contentType: 'application/pdf', size: 12345 },
          ],
        }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({ includeAttachments: true });

      expect(response.results[0].attachments).toHaveLength(1);
      expect(response.results[0].attachments[0]).toMatchObject({
        id: 'att-1',
        name: 'agenda.pdf',
        contentType: 'application/pdf',
        size: 12345,
      });
    });
  });

  describe('search – meeting detection integration', () => {
    it('flags a meeting invite email and sets meetingType to invite', async () => {
      const messages = [
        makeGraphMessage({
          id: 'msg-1',
          subject: 'Invitation: Sprint planning',
          bodyPreview: 'You are invited to the sprint planning session.',
        }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({});

      expect(response.results[0].isMeetingRelated).toBe(true);
      expect(response.results[0].meetingType).toBe('invite');
    });

    it('flags a forwarded meeting invite and sets meetingType to forward', async () => {
      const messages = [
        makeGraphMessage({
          id: 'msg-1',
          subject: 'Fw: Sprint planning invite',
          bodyPreview: 'Join Microsoft Teams Meeting via the link below.',
        }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({});

      expect(response.results[0].isMeetingRelated).toBe(true);
      expect(response.results[0].meetingType).toBe('forward');
    });

    it('sets isMeetingRelated=false for a plain email', async () => {
      const messages = [
        makeGraphMessage({ id: 'msg-1', subject: 'Lunch plans', bodyPreview: 'Let us grab lunch at noon.' }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({});

      expect(response.results[0].isMeetingRelated).toBe(false);
      expect(response.results[0].meetingType).toBeUndefined();
    });
  });

  describe('search – result mapping', () => {
    it('maps sender correctly from the from field', async () => {
      const messages = [
        makeGraphMessage({ from: { emailAddress: { address: 'alice@example.com' } } }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({});

      expect(response.results[0].sender).toBe('alice@example.com');
    });

    it('maps recipients from toRecipients and ccRecipients', async () => {
      const messages = [
        makeGraphMessage({
          toRecipients: [
            { emailAddress: { address: 'bob@example.com' } },
            { emailAddress: { address: 'carol@example.com' } },
          ],
          ccRecipients: [{ emailAddress: { address: 'dave@example.com' } }],
        }),
      ];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({});

      expect(response.results[0].recipients).toEqual([
        'bob@example.com',
        'carol@example.com',
        'dave@example.com',
      ]);
    });

    it('truncates snippet to 200 characters', async () => {
      const longPreview = 'A'.repeat(300);
      const messages = [makeGraphMessage({ bodyPreview: longPreview })];
      const client = makeMockGraphClient(messages);
      const service = new EmailSearchService(client);

      const response = await service.search({});

      expect(response.results[0].snippet).toHaveLength(200);
    });

    it('echoes the original searchParams in the response', async () => {
      const client = makeMockGraphClient([]);
      const service = new EmailSearchService(client);
      const params = { keywords: ['test'], maxResults: 5 };

      const response = await service.search(params);

      expect(response.searchParams).toEqual(params);
    });
  });
});
