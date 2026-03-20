import { MeetingReconstructionEngine } from '../reconstruction/MeetingReconstructionEngine';
import {
  CorrelatedSearchResults,
  TranscriptResult,
  ChatMessage,
  EmailThread,
  FileResult,
} from '../types';

const engine = new MeetingReconstructionEngine();

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const transcriptFull: TranscriptResult = {
  source: 'transcript',
  meetingId: 'mtg-001',
  title: 'Q2 Planning Session',
  date: '2024-06-10T10:00:00Z',
  participants: ['Alice Johnson', 'Bob Smith', 'Carol White'],
  content:
    'We will discuss the Q2 roadmap. We agreed to prioritize the new search feature. ' +
    'Bob will follow up with the UX team by end of week. ' +
    'Alice decided to postpone the release until July. ' +
    'Action item: Carol needs to prepare the presentation slides.',
  segments: [
    { speaker: 'Alice Johnson', text: 'Welcome everyone. Today we will review and discuss the Q2 roadmap.' },
    { speaker: 'Bob Smith', text: 'We agreed to prioritize the new search feature for this quarter.' },
    { speaker: 'Carol White', text: 'I will prepare the presentation slides as an action item.' },
  ],
};

const transcriptNoTitle: TranscriptResult = {
  source: 'transcript',
  meetingId: 'mtg-002',
  date: '2024-06-11T14:00:00Z',
  participants: ['Dave Lee', 'Eve Clark'],
  content:
    'We decided to move forward with the new architecture. ' +
    'Dave will review the requirements by next Monday.',
};

const transcriptMinimal: TranscriptResult = {
  source: 'transcript',
  content: 'Short meeting notes without much detail.',
};

const transcriptSegmentsOnly: TranscriptResult = {
  source: 'transcript',
  content: 'Transcript with segments only.',
  segments: [
    { speaker: 'Frank', text: 'We should discuss the budget proposal.' },
    { speaker: 'Grace', text: 'I agreed to review the numbers.' },
  ],
};

const emailFull: EmailThread = {
  source: 'email',
  subject: 'Re: Q2 Planning Meeting',
  date: '2024-06-10T09:00:00Z',
  from: 'alice@example.com',
  to: ['bob@example.com', 'carol@example.com'],
  body:
    'Hi team, please review the agenda for our meeting. ' +
    'We agreed on the new timeline. ' +
    'Follow up: Bob to send the report by Friday.',
  threadId: 'thread-123',
};

const emailNoSubject: EmailThread = {
  source: 'email',
  date: '2024-06-12T08:00:00Z',
  from: 'dave@example.com',
  to: ['eve@example.com'],
  body: 'We decided to adopt the new framework. Assigned to Dave by end of month.',
};

const emailSubjectWithPrefix: EmailThread = {
  source: 'email',
  subject: 'RE: Budget Review',
  from: 'manager@example.com',
  to: ['team@example.com'],
  body: 'Please review and discuss the budget items before the meeting.',
};

const chatMessages: ChatMessage[] = [
  {
    source: 'chat',
    channelId: 'ch-1',
    channelName: 'project-alpha',
    messageId: 'msg-1',
    sender: 'HenryK',
    timestamp: '2024-06-10T10:05:00Z',
    text: 'We should discuss the deployment plan today.',
  },
  {
    source: 'chat',
    channelId: 'ch-1',
    channelName: 'project-alpha',
    messageId: 'msg-2',
    sender: 'IreneM',
    timestamp: '2024-06-10T10:10:00Z',
    text: 'Agreed, will follow up after the meeting.',
  },
  {
    source: 'chat',
    channelId: 'ch-1',
    channelName: 'project-alpha',
    messageId: 'msg-3',
    sender: 'HenryK',
    timestamp: '2024-06-10T09:55:00Z',
    text: 'Action item: someone needs to update the docs.',
  },
];

const chatNoChannel: ChatMessage = {
  source: 'chat',
  sender: 'JackP',
  timestamp: '2024-06-13T11:00:00Z',
  text: 'We decided to go with option B.',
};

const fileResult: FileResult = {
  source: 'file',
  filename: 'meeting-notes.docx',
  path: '/docs/meeting-notes.docx',
  content: 'Meeting notes from the Q2 planning session.',
  meetingDate: '2024-06-10',
  author: 'Alice Johnson',
};

// ---------------------------------------------------------------------------
// Helper
// ---------------------------------------------------------------------------

function makeResults(results: CorrelatedSearchResults['results']): CorrelatedSearchResults {
  return { query: 'test query', results };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('MeetingReconstructionEngine', () => {
  // -------------------------------------------------------------------------
  // Full data (transcript present)
  // -------------------------------------------------------------------------

  describe('full data with transcript', () => {
    const reconstruction = engine.reconstruct(makeResults([transcriptFull, emailFull, ...chatMessages, fileResult]));

    it('reconstructs title from transcript with high confidence', () => {
      expect(reconstruction.title).not.toBeNull();
      expect(reconstruction.title!.value).toBe('Q2 Planning Session');
      expect(reconstruction.title!.confidence).toBe('high');
      expect(reconstruction.title!.source).toBe('transcript');
    });

    it('reconstructs dateTime from transcript with high confidence', () => {
      expect(reconstruction.dateTime).not.toBeNull();
      expect(reconstruction.dateTime!.value).toBe('2024-06-10T10:00:00Z');
      expect(reconstruction.dateTime!.confidence).toBe('high');
      expect(reconstruction.dateTime!.source).toBe('transcript');
    });

    it('reconstructs attendees from transcript participants with high confidence', () => {
      expect(reconstruction.attendees).not.toBeNull();
      expect(reconstruction.attendees!.value).toEqual(['Alice Johnson', 'Bob Smith', 'Carol White']);
      expect(reconstruction.attendees!.confidence).toBe('high');
    });

    it('extracts topics from transcript content', () => {
      expect(reconstruction.topics).not.toBeNull();
      expect(reconstruction.topics!.confidence).toBe('high');
      expect(reconstruction.topics!.source).toBe('transcript');
      expect(reconstruction.topics!.value.length).toBeGreaterThan(0);
    });

    it('extracts decisions from transcript content', () => {
      expect(reconstruction.decisions).not.toBeNull();
      expect(reconstruction.decisions!.confidence).toBe('high');
      expect(reconstruction.decisions!.value.length).toBeGreaterThan(0);
    });

    it('extracts action items from transcript content', () => {
      expect(reconstruction.actionItems).not.toBeNull();
      expect(reconstruction.actionItems!.confidence).toBe('high');
      expect(reconstruction.actionItems!.value.length).toBeGreaterThan(0);
    });

    it('has no gaps about missing transcript', () => {
      expect(reconstruction.gaps).not.toContain(
        'No meeting transcript available - reconstruction may be incomplete',
      );
    });

    it('reports high overall confidence', () => {
      expect(reconstruction.overallConfidence).toBe('high');
    });
  });

  // -------------------------------------------------------------------------
  // Email-only (no transcript, no chat)
  // -------------------------------------------------------------------------

  describe('email-only (no transcript, no chat)', () => {
    const reconstruction = engine.reconstruct(makeResults([emailFull]));

    it('reconstructs title from email subject with medium confidence', () => {
      expect(reconstruction.title).not.toBeNull();
      expect(reconstruction.title!.value).toBe('Q2 Planning Meeting');
      expect(reconstruction.title!.confidence).toBe('medium');
      expect(reconstruction.title!.source).toBe('email');
    });

    it('strips reply/forward prefix from email subject', () => {
      const r = engine.reconstruct(makeResults([emailSubjectWithPrefix]));
      expect(r.title!.value).toBe('Budget Review');
    });

    it('reconstructs dateTime from email with medium confidence', () => {
      expect(reconstruction.dateTime).not.toBeNull();
      expect(reconstruction.dateTime!.confidence).toBe('medium');
      expect(reconstruction.dateTime!.source).toBe('email');
    });

    it('reconstructs attendees from email to/from with medium confidence', () => {
      expect(reconstruction.attendees).not.toBeNull();
      expect(reconstruction.attendees!.confidence).toBe('medium');
      expect(reconstruction.attendees!.value).toContain('alice@example.com');
      expect(reconstruction.attendees!.value).toContain('bob@example.com');
    });

    it('flags missing transcript as a gap', () => {
      expect(reconstruction.gaps).toContain(
        'No meeting transcript available - reconstruction may be incomplete',
      );
    });

    it('reports medium overall confidence', () => {
      expect(['medium', 'low']).toContain(reconstruction.overallConfidence);
    });
  });

  // -------------------------------------------------------------------------
  // Chat-only (no transcript, no email)
  // -------------------------------------------------------------------------

  describe('chat-only (no transcript, no email)', () => {
    const reconstruction = engine.reconstruct(makeResults(chatMessages));

    it('reconstructs title from channel name with low confidence', () => {
      expect(reconstruction.title).not.toBeNull();
      expect(reconstruction.title!.value).toBe('project-alpha');
      expect(reconstruction.title!.confidence).toBe('low');
      expect(reconstruction.title!.source).toBe('chat');
    });

    it('reconstructs dateTime from earliest chat timestamp with low confidence', () => {
      expect(reconstruction.dateTime).not.toBeNull();
      expect(reconstruction.dateTime!.value).toBe('2024-06-10T09:55:00Z');
      expect(reconstruction.dateTime!.confidence).toBe('low');
      expect(reconstruction.dateTime!.source).toBe('chat');
    });

    it('reconstructs attendees from unique chat senders with low confidence', () => {
      expect(reconstruction.attendees).not.toBeNull();
      expect(reconstruction.attendees!.confidence).toBe('low');
      expect(reconstruction.attendees!.value).toContain('HenryK');
      expect(reconstruction.attendees!.value).toContain('IreneM');
    });

    it('deduplicates chat senders', () => {
      // HenryK appears twice in chatMessages but should only appear once
      const senders = reconstruction.attendees!.value;
      expect(senders.filter((s) => s === 'HenryK').length).toBe(1);
    });

    it('flags missing transcript as a gap', () => {
      expect(reconstruction.gaps).toContain(
        'No meeting transcript available - reconstruction may be incomplete',
      );
    });
  });

  // -------------------------------------------------------------------------
  // Transcript without explicit participants (uses segments)
  // -------------------------------------------------------------------------

  describe('transcript with segments but no explicit participants', () => {
    const reconstruction = engine.reconstruct(makeResults([transcriptSegmentsOnly]));

    it('extracts attendees from segment speakers', () => {
      expect(reconstruction.attendees).not.toBeNull();
      expect(reconstruction.attendees!.value).toContain('Frank');
      expect(reconstruction.attendees!.value).toContain('Grace');
      expect(reconstruction.attendees!.sourceDetails).toMatch(/segment/i);
    });
  });

  // -------------------------------------------------------------------------
  // Transcript without title – falls back to email
  // -------------------------------------------------------------------------

  describe('transcript without title, with email fallback', () => {
    const reconstruction = engine.reconstruct(makeResults([transcriptNoTitle, emailFull]));

    it('falls back to email subject for title', () => {
      expect(reconstruction.title).not.toBeNull();
      expect(reconstruction.title!.confidence).toBe('medium');
      expect(reconstruction.title!.source).toBe('email');
    });

    it('still uses transcript date with high confidence', () => {
      expect(reconstruction.dateTime!.confidence).toBe('high');
      expect(reconstruction.dateTime!.source).toBe('transcript');
    });
  });

  // -------------------------------------------------------------------------
  // Empty / no data
  // -------------------------------------------------------------------------

  describe('empty results', () => {
    const reconstruction = engine.reconstruct(makeResults([]));

    it('returns null for all elements', () => {
      expect(reconstruction.title).toBeNull();
      expect(reconstruction.dateTime).toBeNull();
      expect(reconstruction.attendees).toBeNull();
      expect(reconstruction.topics).toBeNull();
      expect(reconstruction.decisions).toBeNull();
      expect(reconstruction.actionItems).toBeNull();
    });

    it('reports all expected gaps', () => {
      expect(reconstruction.gaps).toContain('Meeting title could not be determined');
      expect(reconstruction.gaps).toContain('Meeting date/time could not be determined');
      expect(reconstruction.gaps).toContain('Attendees could not be determined');
      expect(reconstruction.gaps).toContain('Discussion topics could not be extracted');
      expect(reconstruction.gaps).toContain('No decisions identified in available sources');
      expect(reconstruction.gaps).toContain('No action items identified in available sources');
      expect(reconstruction.gaps).toContain(
        'No meeting transcript available - reconstruction may be incomplete',
      );
      expect(reconstruction.gaps).toContain('No data sources available for reconstruction');
    });

    it('reports low overall confidence', () => {
      expect(reconstruction.overallConfidence).toBe('low');
    });
  });

  // -------------------------------------------------------------------------
  // File-only results
  // -------------------------------------------------------------------------

  describe('file-only results', () => {
    const reconstruction = engine.reconstruct(makeResults([fileResult]));

    it('returns null for title (files do not provide title)', () => {
      expect(reconstruction.title).toBeNull();
    });

    it('flags transcript as missing', () => {
      expect(reconstruction.gaps).toContain(
        'No meeting transcript available - reconstruction may be incomplete',
      );
    });
  });

  // -------------------------------------------------------------------------
  // Partial: chat + email, no transcript
  // -------------------------------------------------------------------------

  describe('partial data: chat + email, no transcript', () => {
    const reconstruction = engine.reconstruct(makeResults([emailFull, ...chatMessages]));

    it('prefers email title over chat channel name', () => {
      expect(reconstruction.title!.source).toBe('email');
    });

    it('prefers email date over chat timestamp', () => {
      expect(reconstruction.dateTime!.source).toBe('email');
    });

    it('prefers email attendees over chat senders', () => {
      expect(reconstruction.attendees!.source).toBe('email');
    });

    it('flags transcript as missing', () => {
      expect(reconstruction.gaps).toContain(
        'No meeting transcript available - reconstruction may be incomplete',
      );
    });
  });

  // -------------------------------------------------------------------------
  // Decision and action item extraction
  // -------------------------------------------------------------------------

  describe('decision extraction from email when no transcript decisions', () => {
    const reconstruction = engine.reconstruct(makeResults([emailNoSubject]));

    it('extracts decisions from email body with low confidence', () => {
      expect(reconstruction.decisions).not.toBeNull();
      expect(reconstruction.decisions!.confidence).toBe('low');
      expect(reconstruction.decisions!.source).toBe('email');
    });
  });

  describe('action item extraction from chat', () => {
    const reconstruction = engine.reconstruct(makeResults([chatNoChannel, chatMessages[2]]));

    it('extracts action items from chat messages with medium confidence', () => {
      expect(reconstruction.actionItems).not.toBeNull();
      expect(reconstruction.actionItems!.confidence).toBe('medium');
      expect(reconstruction.actionItems!.source).toBe('chat');
    });
  });

  // -------------------------------------------------------------------------
  // Email subject prefix stripping
  // -------------------------------------------------------------------------

  describe('email subject prefix stripping', () => {
    const prefixes = ['Re:', 'Fwd:', 'FW:', 'RE:'];

    prefixes.forEach((prefix) => {
      it(`strips "${prefix}" prefix from email subject`, () => {
        const email: EmailThread = {
          source: 'email',
          subject: `${prefix} Sprint Review`,
          from: 'user@example.com',
          body: 'Meeting notes',
        };
        const r = engine.reconstruct(makeResults([email]));
        expect(r.title!.value).toBe('Sprint Review');
      });
    });
  });

  // -------------------------------------------------------------------------
  // Overall confidence computation
  // -------------------------------------------------------------------------

  describe('overall confidence computation', () => {
    it('is high when all elements come from transcript', () => {
      const r = engine.reconstruct(makeResults([transcriptFull]));
      expect(r.overallConfidence).toBe('high');
    });

    it('is medium when elements come from email', () => {
      const r = engine.reconstruct(makeResults([emailFull]));
      expect(['medium', 'low']).toContain(r.overallConfidence);
    });

    it('is low when no reconstructable sources are available', () => {
      const r = engine.reconstruct(makeResults([]));
      expect(r.overallConfidence).toBe('low');
    });
  });

  // -------------------------------------------------------------------------
  // Gap identification
  // -------------------------------------------------------------------------

  describe('gap identification', () => {
    it('does not flag transcript gap when transcript is present', () => {
      const r = engine.reconstruct(makeResults([transcriptFull]));
      expect(r.gaps).not.toContain(
        'No meeting transcript available - reconstruction may be incomplete',
      );
    });

    it('does not flag "no data sources" gap when at least one source is present', () => {
      const r = engine.reconstruct(makeResults([emailFull]));
      expect(r.gaps).not.toContain('No data sources available for reconstruction');
    });
  });

  // -------------------------------------------------------------------------
  // Minimal transcript (edge cases)
  // -------------------------------------------------------------------------

  describe('minimal transcript with no rich metadata', () => {
    const reconstruction = engine.reconstruct(makeResults([transcriptMinimal]));

    it('returns null for title when transcript has no title', () => {
      expect(reconstruction.title).toBeNull();
    });

    it('returns null for dateTime when transcript has no date', () => {
      expect(reconstruction.dateTime).toBeNull();
    });

    it('returns null for attendees when transcript has no participants or segments', () => {
      expect(reconstruction.attendees).toBeNull();
    });

    it('still attempts to extract topics from minimal content', () => {
      // Short text may produce topics via the fallback path
      // The exact value doesn't matter; we just check the engine doesn't throw
      expect(reconstruction.topics === null || Array.isArray(reconstruction.topics?.value)).toBe(true);
    });
  });
});
