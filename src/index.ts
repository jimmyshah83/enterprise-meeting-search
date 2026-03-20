export { EmailSearchService } from './email/emailSearchService';
export { createGraphClientWrapper, GraphClientWrapper } from './graph/graphClient';
export type {
  EmailAttachment,
  EmailResult,
  EmailSearchParams,
  EmailSearchResponse,
  GraphAttachment,
  GraphMessage,
  MeetingEmailType,
} from './email/emailSearchTypes';
export type { MeetingDetectionResult } from './email/meetingDetector';
export { detectMeetingEmail } from './email/meetingDetector';
