import { detectMeetingEmail } from '../../src/email/meetingDetector';
import { GraphMessage } from '../../src/email/emailSearchTypes';

function makeMessage(overrides: Partial<GraphMessage> = {}): GraphMessage {
  return {
    id: 'msg-1',
    subject: '',
    bodyPreview: '',
    receivedDateTime: '2024-01-15T10:00:00Z',
    from: { emailAddress: { address: 'sender@example.com' } },
    toRecipients: [],
    categories: [],
    ...overrides,
  };
}

describe('detectMeetingEmail', () => {
  describe('non-meeting emails', () => {
    it('returns isMeetingRelated=false for a plain email', () => {
      const result = detectMeetingEmail(makeMessage({ subject: 'Hello there', bodyPreview: 'Just checking in.' }));
      expect(result.isMeetingRelated).toBe(false);
      expect(result.meetingType).toBeUndefined();
    });

    it('returns isMeetingRelated=false for an empty message', () => {
      const result = detectMeetingEmail(makeMessage());
      expect(result.isMeetingRelated).toBe(false);
    });
  });

  describe('meeting invites', () => {
    it('detects a subject containing "invite"', () => {
      const result = detectMeetingEmail(makeMessage({ subject: 'Team sync invite' }));
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('invite');
    });

    it('detects a subject containing "Invitation"', () => {
      const result = detectMeetingEmail(makeMessage({ subject: 'Invitation: Q1 planning' }));
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('invite');
    });

    it('detects a subject containing "meeting request"', () => {
      const result = detectMeetingEmail(makeMessage({ subject: 'Meeting Request: 1:1 with Alice' }));
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('invite');
    });

    it('detects meeting category tag', () => {
      const result = detectMeetingEmail(makeMessage({ categories: ['Meeting'] }));
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('invite');
    });
  });

  describe('forwarded meeting invites', () => {
    it('detects a forwarded meeting invite via subject prefix and body hint', () => {
      const result = detectMeetingEmail(
        makeMessage({
          subject: 'Fw: Team sync invite',
          bodyPreview: 'Please join Microsoft Teams Meeting at the link below.',
        }),
      );
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('forward');
    });

    it('detects FWD: prefix as a forward', () => {
      const result = detectMeetingEmail(
        makeMessage({
          subject: 'FWD: Schedule a meeting with the client',
          bodyPreview: 'Join Microsoft Teams Meeting',
        }),
      );
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('forward');
    });
  });

  describe('meeting replies', () => {
    it('detects a reply to a meeting invite with Teams body hint', () => {
      const result = detectMeetingEmail(
        makeMessage({
          subject: 'Re: Weekly standup invite',
          bodyPreview: 'Join Microsoft Teams Meeting at the following link.',
        }),
      );
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('reply');
    });
  });

  describe('meeting cancellations', () => {
    it('detects "Canceled" in the subject', () => {
      const result = detectMeetingEmail(makeMessage({ subject: 'Canceled: Weekly sync' }));
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('cancellation');
    });

    it('detects "Cancelled" (UK spelling) in the subject', () => {
      const result = detectMeetingEmail(makeMessage({ subject: 'Cancelled: Q2 review' }));
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('cancellation');
    });

    it('detects "cancellation" in the subject', () => {
      const result = detectMeetingEmail(makeMessage({ subject: 'Meeting cancellation notice' }));
      expect(result.isMeetingRelated).toBe(true);
      expect(result.meetingType).toBe('cancellation');
    });
  });

  describe('body-hint detection', () => {
    it('detects Join Zoom Meeting in the body', () => {
      const result = detectMeetingEmail(
        makeMessage({ subject: 'Action required', bodyPreview: 'Please join Zoom Meeting using the link.' }),
      );
      expect(result.isMeetingRelated).toBe(true);
    });

    it('detects Join Google Meet in the body', () => {
      const result = detectMeetingEmail(
        makeMessage({ subject: 'Discussion', bodyPreview: 'Click here to join Google Meet.' }),
      );
      expect(result.isMeetingRelated).toBe(true);
    });

    it('detects meeting link in the body', () => {
      const result = detectMeetingEmail(
        makeMessage({ subject: 'Project update', bodyPreview: 'The meeting link is attached below.' }),
      );
      expect(result.isMeetingRelated).toBe(true);
    });

    it('detects dial-in number in the body', () => {
      const result = detectMeetingEmail(
        makeMessage({ subject: 'Call details', bodyPreview: 'Use the dial-in number: +1 800 555 0000.' }),
      );
      expect(result.isMeetingRelated).toBe(true);
    });
  });
});
