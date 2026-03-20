/**
 * Meeting reconstruction engine.
 *
 * Reconstructs a structured meeting summary from a UnifiedMeetingResult,
 * extracting titles, date/time, attendees, topics, decisions and action items.
 */

import {
  UnifiedMeetingResult,
  MeetingReconstruction,
  ReconstructedElement,
  ConfidenceLevel,
  ReconstructionSourceType,
  ActionItem,
} from '../types';

const DECISION_KEYWORDS = [
  'decided', 'agreed', 'approved', 'resolved', 'concluded',
  'confirmed', 'accepted', 'rejected', 'will proceed', 'we will',
];

const ACTION_ITEM_KEYWORDS = [
  'action item', 'todo', 'to do', 'follow up', 'follow-up',
  'will', 'should', 'needs to', 'assigned to', 'responsible for',
  'by next', 'by end of', 'deadline',
];

const TOPIC_KEYWORDS = [
  'discuss', 'agenda', 'topic', 'review', 'present',
  'cover', 'address', 'consider',
];

function elem<T>(value: T, confidence: ConfidenceLevel, source: ReconstructionSourceType, sourceDetails?: string): ReconstructedElement<T> {
  return { value, confidence, source, sourceDetails };
}

function extractSentencesContaining(text: string, keywords: string[]): string[] {
  const sentences = text.split(/[.!?](?:\s+|$)/);
  return sentences.filter((s) =>
    keywords.some((kw) => s.toLowerCase().includes(kw.toLowerCase())),
  );
}

export class MeetingReconstructionEngine {
  reconstruct(meetingResult: UnifiedMeetingResult): MeetingReconstruction {
    const { meetingId, calendarEvent, transcripts, chats, emails } = meetingResult;

    const title = this.reconstructTitle(calendarEvent, transcripts.map(t => t.content), emails.map(e => e.subject));
    const dateTime = this.reconstructDateTime(calendarEvent, transcripts.map(t => t.timestamp.toISOString()), emails.map(e => e.timestamp.toISOString()));
    const attendees = this.reconstructAttendees(calendarEvent, transcripts.flatMap(t => t.speakers), emails.flatMap(e => [e.from, ...e.to]));
    const topics = this.reconstructTopics(transcripts.map(t => t.content), chats.map(c => c.content), emails.map(e => e.body));
    const decisions = this.reconstructDecisions(transcripts.map(t => t.content), chats.map(c => c.content), emails.map(e => e.body));
    const actionItems = this.reconstructActionItems(transcripts.map(t => t.content), chats.map(c => c.content), emails.map(e => e.body));

    const gaps = this.identifyGaps({ title, dateTime, attendees, topics, decisions, actionItems }, transcripts.length > 0, chats.length > 0, emails.length > 0);
    const overallConfidence = this.computeOverallConfidence([title, dateTime, attendees, topics, decisions, actionItems]);

    return {
      meetingId,
      title,
      dateTime,
      attendees,
      topics,
      decisions,
      actionItems,
      overallConfidence,
      gaps,
      reconstructedAt: new Date(),
    };
  }

  private reconstructTitle(
    calendarEvent: UnifiedMeetingResult['calendarEvent'],
    transcriptContents: string[],
    emailSubjects: string[],
  ): ReconstructedElement<string> | null {
    if (calendarEvent?.title) {
      return elem(calendarEvent.title, 'high', 'calendar', 'Calendar event title');
    }
    const emailSubject = emailSubjects.find((s) => s && s.trim().length > 0);
    if (emailSubject) {
      return elem(emailSubject, 'medium', 'email', 'Email subject line');
    }
    const transcriptTitle = transcriptContents.find((c) => c && c.length > 0);
    if (transcriptTitle) {
      const firstLine = transcriptTitle.split('\n')[0].slice(0, 100);
      return elem(firstLine, 'low', 'transcript', 'First line of transcript');
    }
    return null;
  }

  private reconstructDateTime(
    calendarEvent: UnifiedMeetingResult['calendarEvent'],
    transcriptTimestamps: string[],
    emailTimestamps: string[],
  ): ReconstructedElement<string> | null {
    if (calendarEvent?.startTime) {
      return elem(calendarEvent.startTime.toISOString(), 'high', 'calendar', 'Calendar event start time');
    }
    const ts = transcriptTimestamps.find((t) => t && t.length > 0) ?? emailTimestamps.find((t) => t && t.length > 0);
    if (ts) {
      return elem(ts, 'medium', transcriptTimestamps.length > 0 ? 'transcript' : 'email', 'Timestamp from source');
    }
    return null;
  }

  private reconstructAttendees(
    calendarEvent: UnifiedMeetingResult['calendarEvent'],
    transcriptSpeakers: string[],
    emailParticipants: string[],
  ): ReconstructedElement<string[]> | null {
    if (calendarEvent?.attendees && calendarEvent.attendees.length > 0) {
      return elem([...new Set(calendarEvent.attendees)], 'high', 'calendar', 'Calendar event attendee list');
    }
    const combined = [...new Set([...transcriptSpeakers, ...emailParticipants].filter(Boolean))];
    if (combined.length > 0) {
      const source: ReconstructionSourceType = transcriptSpeakers.length > 0 ? 'transcript' : 'email';
      return elem(combined, transcriptSpeakers.length > 0 ? 'medium' : 'low', source, 'Derived from speakers/participants');
    }
    return null;
  }

  private reconstructTopics(
    transcriptContents: string[],
    chatContents: string[],
    emailBodies: string[],
  ): ReconstructedElement<string[]> | null {
    const allContent = [...transcriptContents, ...chatContents, ...emailBodies];
    const topicSentences: string[] = [];
    for (const content of allContent) {
      topicSentences.push(...extractSentencesContaining(content, TOPIC_KEYWORDS));
    }
    if (topicSentences.length === 0) return null;
    const source: ReconstructionSourceType = transcriptContents.length > 0 ? 'transcript' : chatContents.length > 0 ? 'chat' : 'email';
    const confidence: ConfidenceLevel = transcriptContents.length > 0 ? 'high' : chatContents.length > 0 ? 'medium' : 'low';
    return elem(topicSentences.slice(0, 5), confidence, source, 'Extracted from meeting content');
  }

  private reconstructDecisions(
    transcriptContents: string[],
    chatContents: string[],
    emailBodies: string[],
  ): ReconstructedElement<string[]> | null {
    const allContent = [...transcriptContents, ...chatContents, ...emailBodies];
    const decisionSentences: string[] = [];
    for (const content of allContent) {
      decisionSentences.push(...extractSentencesContaining(content, DECISION_KEYWORDS));
    }
    if (decisionSentences.length === 0) return null;
    const source: ReconstructionSourceType = transcriptContents.length > 0 ? 'transcript' : chatContents.length > 0 ? 'chat' : 'email';
    const confidence: ConfidenceLevel = transcriptContents.length > 0 ? 'high' : chatContents.length > 0 ? 'medium' : 'low';
    return elem(decisionSentences.slice(0, 5), confidence, source, 'Extracted from meeting content');
  }

  private reconstructActionItems(
    transcriptContents: string[],
    chatContents: string[],
    emailBodies: string[],
  ): ReconstructedElement<ActionItem[]> | null {
    const allContent = [...transcriptContents, ...chatContents, ...emailBodies];
    const actionSentences: string[] = [];
    for (const content of allContent) {
      actionSentences.push(...extractSentencesContaining(content, ACTION_ITEM_KEYWORDS));
    }
    if (actionSentences.length === 0) return null;
    const items: ActionItem[] = actionSentences.slice(0, 5).map((s) => ({ description: s.trim() }));
    const source: ReconstructionSourceType = transcriptContents.length > 0 ? 'transcript' : chatContents.length > 0 ? 'chat' : 'email';
    const confidence: ConfidenceLevel = transcriptContents.length > 0 ? 'high' : chatContents.length > 0 ? 'medium' : 'low';
    return elem(items, confidence, source, 'Extracted from meeting content');
  }

  private identifyGaps(
    elements: Record<string, ReconstructedElement<unknown> | null>,
    hasTranscripts: boolean,
    hasChats: boolean,
    hasEmails: boolean,
  ): string[] {
    const gaps: string[] = [];
    for (const [key, element] of Object.entries(elements)) {
      if (element === null) gaps.push(`Could not reconstruct: ${key}`);
    }
    if (!hasTranscripts) gaps.push('No transcript data available — reconstruction confidence is reduced');
    if (!hasChats && !hasEmails) gaps.push('No chat or email data available');
    return gaps;
  }

  private computeOverallConfidence(
    elements: Array<ReconstructedElement<unknown> | null>,
  ): ConfidenceLevel {
    const present = elements.filter(Boolean) as ReconstructedElement<unknown>[];
    if (present.length === 0) return 'inferred';
    const levels: ConfidenceLevel[] = ['high', 'medium', 'low', 'inferred'];
    const worstIndex = Math.max(...present.map((e) => levels.indexOf(e.confidence)));
    const avgIndex = Math.round(present.reduce((sum, e) => sum + levels.indexOf(e.confidence), 0) / present.length);
    const dominantIndex = Math.max(worstIndex - 1, avgIndex);
    return levels[Math.min(dominantIndex, levels.length - 1)];
  }
}
