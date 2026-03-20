/**
 * Entry point – re-exports the public API surface so consumers can simply:
 *   import { search } from 'enterprise-meeting-search';
 */

export { search, DomainAdapters } from './orchestrator';
export { buildDiagnostics, formatDiagnosticText } from './diagnostics';
export * from './types';
export { searchCalendar } from './domains/calendar';
export { searchTranscripts } from './domains/transcripts';
export { searchChats } from './domains/chats';
export { searchEmail } from './domains/email';
export { searchFiles } from './domains/files';
