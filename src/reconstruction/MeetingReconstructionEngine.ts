import {
  CorrelatedSearchResults,
  TranscriptResult,
  ChatMessage,
  EmailThread,
  ConfidenceLevel,
  ReconstructedElement,
  MeetingReconstruction,
  SearchResult,
} from '../types';

/** Keywords that typically signal a decision was made. */
const DECISION_KEYWORDS: string[] = [
  'decided',
  'agreed',
  'approved',
  'resolved',
  'concluded',
  'confirmed',
  'accepted',
  'rejected',
  'will proceed',
  'we will',
];

/** Keywords that typically signal an action item was created. */
const ACTION_ITEM_KEYWORDS: string[] = [
  'action item',
  'todo',
  'to do',
  'follow up',
  'follow-up',
  'will',
  'should',
  'needs to',
  'assigned to',
  'responsible for',
  'by next',
  'by end of',
  'deadline',
];

/** Keywords that typically introduce a discussion topic. */
const TOPIC_KEYWORDS: string[] = [
  'discuss',
  'agenda',
  'topic',
  'review',
  'present',
  'cover',
  'address',
  'consider',
];

/**
 * Reconstructs a structured meeting summary from correlated cross-domain
 * search results (transcripts, chat messages, email threads, and files).
 *
 * Reconstruction strategy (in priority order for each element):
 *  1. Meeting transcript  → `high` confidence
 *  2. Email thread        → `medium` confidence
 *  3. Chat messages       → `low` confidence
 *
 * Elements that cannot be determined are left as `null` and a corresponding
 * entry is added to the `gaps` array so callers know what is missing.
 */
export class MeetingReconstructionEngine {
  /**
   * Reconstruct a meeting from the given correlated search results.
   *
   * @param correlatedResults - The output of the cross-domain search orchestrator.
   * @returns A {@link MeetingReconstruction} with confidence and source metadata
   *          for every element, and an explicit list of information gaps.
   */
  reconstruct(correlatedResults: CorrelatedSearchResults): MeetingReconstruction {
    const { results } = correlatedResults;

    const transcripts = results.filter((r): r is TranscriptResult => r.source === 'transcript');
    const chatMessages = results.filter((r): r is ChatMessage => r.source === 'chat');
    const emails = results.filter((r): r is EmailThread => r.source === 'email');

    const title = this.reconstructTitle(transcripts, emails, chatMessages);
    const dateTime = this.reconstructDateTime(transcripts, emails, chatMessages);
    const attendees = this.reconstructAttendees(transcripts, emails, chatMessages);
    const topics = this.reconstructTopics(transcripts, chatMessages, emails);
    const decisions = this.reconstructDecisions(transcripts, chatMessages, emails);
    const actionItems = this.reconstructActionItems(transcripts, chatMessages, emails);

    const gaps = this.identifyGaps(
      { title, dateTime, attendees, topics, decisions, actionItems },
      transcripts.length > 0,
      chatMessages.length > 0,
      emails.length > 0,
    );

    const overallConfidence = this.computeOverallConfidence([
      title,
      dateTime,
      attendees,
      topics,
      decisions,
      actionItems,
    ]);

    return { title, dateTime, attendees, topics, decisions, actionItems, gaps, overallConfidence };
  }

  // ---------------------------------------------------------------------------
  // Private reconstruction helpers
  // ---------------------------------------------------------------------------

  private reconstructTitle(
    transcripts: TranscriptResult[],
    emails: EmailThread[],
    chats: ChatMessage[],
  ): ReconstructedElement<string> | null {
    for (const t of transcripts) {
      if (t.title) {
        return { value: t.title, confidence: 'high', source: 'transcript' };
      }
    }

    for (const e of emails) {
      if (e.subject) {
        const cleaned = e.subject.replace(/^(Re:|Fwd:|FW:|RE:)\s*/gi, '').trim();
        return { value: cleaned, confidence: 'medium', source: 'email', sourceDetails: e.subject };
      }
    }

    for (const c of chats) {
      if (c.channelName) {
        return { value: c.channelName, confidence: 'low', source: 'chat' };
      }
    }

    return null;
  }

  private reconstructDateTime(
    transcripts: TranscriptResult[],
    emails: EmailThread[],
    chats: ChatMessage[],
  ): ReconstructedElement<string> | null {
    for (const t of transcripts) {
      if (t.date) {
        return { value: t.date, confidence: 'high', source: 'transcript' };
      }
    }

    for (const e of emails) {
      if (e.date) {
        return { value: e.date, confidence: 'medium', source: 'email' };
      }
    }

    if (chats.length > 0) {
      const sorted = [...chats].sort((a, b) => a.timestamp.localeCompare(b.timestamp));
      return {
        value: sorted[0].timestamp,
        confidence: 'low',
        source: 'chat',
        sourceDetails: 'earliest chat message timestamp',
      };
    }

    return null;
  }

  private reconstructAttendees(
    transcripts: TranscriptResult[],
    emails: EmailThread[],
    chats: ChatMessage[],
  ): ReconstructedElement<string[]> | null {
    for (const t of transcripts) {
      if (t.participants && t.participants.length > 0) {
        return { value: t.participants, confidence: 'high', source: 'transcript' };
      }
      if (t.segments && t.segments.length > 0) {
        const speakers = [...new Set(t.segments.map((s) => s.speaker).filter(Boolean))];
        if (speakers.length > 0) {
          return {
            value: speakers,
            confidence: 'high',
            source: 'transcript',
            sourceDetails: 'extracted from transcript segments',
          };
        }
      }
    }

    const emailParticipants: string[] = [];
    for (const e of emails) {
      if (e.from) emailParticipants.push(e.from);
      if (e.to) emailParticipants.push(...e.to);
    }
    if (emailParticipants.length > 0) {
      return { value: [...new Set(emailParticipants)], confidence: 'medium', source: 'email' };
    }

    if (chats.length > 0) {
      const senders = [...new Set(chats.map((c) => c.sender).filter(Boolean))];
      if (senders.length > 0) {
        return {
          value: senders,
          confidence: 'low',
          source: 'chat',
          sourceDetails: 'chat message senders',
        };
      }
    }

    return null;
  }

  private reconstructTopics(
    transcripts: TranscriptResult[],
    chats: ChatMessage[],
    emails: EmailThread[],
  ): ReconstructedElement<string[]> | null {
    if (transcripts.length > 0) {
      const text = transcripts.map((t) => t.content).join('\n');
      const topics = this.extractTopicsFromText(text);
      if (topics.length > 0) {
        return { value: topics, confidence: 'high', source: 'transcript' };
      }
    }

    if (chats.length > 0) {
      const text = chats.map((c) => c.text).join('\n');
      const topics = this.extractTopicsFromText(text);
      if (topics.length > 0) {
        return { value: topics, confidence: 'medium', source: 'chat' };
      }
    }

    if (emails.length > 0) {
      const text = emails.map((e) => e.body).join('\n');
      const topics = this.extractTopicsFromText(text);
      if (topics.length > 0) {
        return { value: topics, confidence: 'low', source: 'email' };
      }
    }

    return null;
  }

  private reconstructDecisions(
    transcripts: TranscriptResult[],
    chats: ChatMessage[],
    emails: EmailThread[],
  ): ReconstructedElement<string[]> | null {
    if (transcripts.length > 0) {
      const text = transcripts.map((t) => t.content).join('\n');
      const decisions = this.extractSentencesWithKeywords(text, DECISION_KEYWORDS);
      if (decisions.length > 0) {
        return { value: decisions, confidence: 'high', source: 'transcript' };
      }
    }

    if (chats.length > 0) {
      const text = chats.map((c) => c.text).join('\n');
      const decisions = this.extractSentencesWithKeywords(text, DECISION_KEYWORDS);
      if (decisions.length > 0) {
        return { value: decisions, confidence: 'medium', source: 'chat' };
      }
    }

    if (emails.length > 0) {
      const text = emails.map((e) => e.body).join('\n');
      const decisions = this.extractSentencesWithKeywords(text, DECISION_KEYWORDS);
      if (decisions.length > 0) {
        return { value: decisions, confidence: 'low', source: 'email' };
      }
    }

    return null;
  }

  private reconstructActionItems(
    transcripts: TranscriptResult[],
    chats: ChatMessage[],
    emails: EmailThread[],
  ): ReconstructedElement<string[]> | null {
    if (transcripts.length > 0) {
      const text = transcripts.map((t) => t.content).join('\n');
      const items = this.extractSentencesWithKeywords(text, ACTION_ITEM_KEYWORDS);
      if (items.length > 0) {
        return { value: items, confidence: 'high', source: 'transcript' };
      }
    }

    if (chats.length > 0) {
      const text = chats.map((c) => c.text).join('\n');
      const items = this.extractSentencesWithKeywords(text, ACTION_ITEM_KEYWORDS);
      if (items.length > 0) {
        return { value: items, confidence: 'medium', source: 'chat' };
      }
    }

    if (emails.length > 0) {
      const text = emails.map((e) => e.body).join('\n');
      const items = this.extractSentencesWithKeywords(text, ACTION_ITEM_KEYWORDS);
      if (items.length > 0) {
        return { value: items, confidence: 'low', source: 'email' };
      }
    }

    return null;
  }

  // ---------------------------------------------------------------------------
  // Private text analysis utilities
  // ---------------------------------------------------------------------------

  /**
   * Extracts likely discussion topic sentences from free text.
   * First tries to find sentences with explicit topic-signal keywords;
   * if none are found it falls back to the first few non-trivial sentences
   * so that the engine always produces some output when content exists.
   */
  private extractTopicsFromText(text: string): string[] {
    const sentences = text
      .split(/[.!?]\s+/)
      .map((s) => s.trim())
      .filter((s) => s.length > 10);

    const matched = sentences.filter((s) =>
      TOPIC_KEYWORDS.some((kw) => s.toLowerCase().includes(kw)),
    );

    const candidates = matched.length > 0 ? matched : sentences.slice(0, 3);
    return candidates.slice(0, 5).map((s) => (s.length > 150 ? `${s.substring(0, 150)}...` : s));
  }

  /**
   * Extracts sentences that contain at least one of the supplied keywords.
   * Returns up to five results, each capped at 200 characters.
   */
  private extractSentencesWithKeywords(text: string, keywords: string[]): string[] {
    const sentences = text
      .split(/[.!?\n]+/)
      .map((s) => s.trim())
      .filter((s) => s.length > 10);

    return sentences
      .filter((s) => keywords.some((kw) => s.toLowerCase().includes(kw)))
      .slice(0, 5)
      .map((s) => (s.length > 200 ? `${s.substring(0, 200)}...` : s));
  }

  /**
   * Identifies information that could not be reconstructed and explains why.
   */
  private identifyGaps(
    elements: {
      title: ReconstructedElement<string> | null;
      dateTime: ReconstructedElement<string> | null;
      attendees: ReconstructedElement<string[]> | null;
      topics: ReconstructedElement<string[]> | null;
      decisions: ReconstructedElement<string[]> | null;
      actionItems: ReconstructedElement<string[]> | null;
    },
    hasTranscript: boolean,
    hasChat: boolean,
    hasEmail: boolean,
  ): string[] {
    const gaps: string[] = [];

    if (!elements.title) gaps.push('Meeting title could not be determined');
    if (!elements.dateTime) gaps.push('Meeting date/time could not be determined');
    if (!elements.attendees) gaps.push('Attendees could not be determined');
    if (!elements.topics) gaps.push('Discussion topics could not be extracted');
    if (!elements.decisions) gaps.push('No decisions identified in available sources');
    if (!elements.actionItems) gaps.push('No action items identified in available sources');

    if (!hasTranscript) {
      gaps.push('No meeting transcript available - reconstruction may be incomplete');
    }
    if (!hasChat && !hasEmail && !hasTranscript) {
      gaps.push('No data sources available for reconstruction');
    }

    return gaps;
  }

  /**
   * Computes a single aggregate confidence level across all reconstructed
   * elements, weighted by how many elements were successfully reconstructed.
   */
  private computeOverallConfidence(
    elements: Array<ReconstructedElement<unknown> | null>,
  ): ConfidenceLevel {
    const present = elements.filter((e): e is ReconstructedElement<unknown> => e !== null);
    if (present.length === 0) return 'low';

    const confidenceScore: Record<ConfidenceLevel, number> = {
      high: 3,
      medium: 2,
      low: 1,
      inferred: 0,
    };

    const total = present.reduce((sum, e) => sum + confidenceScore[e.confidence], 0);
    const avg = total / present.length;

    if (avg >= 2.5) return 'high';
    if (avg >= 1.5) return 'medium';
    if (avg >= 0.5) return 'low';
    return 'inferred';
  }
}

// Convenience re-export of types so consumers can import from one place.
export type { SearchResult, CorrelatedSearchResults, MeetingReconstruction, ReconstructedElement, ConfidenceLevel, SourceType } from '../types';
