import { parseMeetingType } from './meetingTypeParser';

describe('parseMeetingType', () => {
  test('"Teams meeting" → teams', () => {
    expect(parseMeetingType('Teams meeting')).toBe('teams');
  });

  test('"Microsoft Teams" → teams', () => {
    expect(parseMeetingType('Microsoft Teams call')).toBe('teams');
  });

  test('"MS Teams" → teams', () => {
    expect(parseMeetingType('MS Teams')).toBe('teams');
  });

  test('"Teams" alone → teams', () => {
    expect(parseMeetingType('It was a Teams call')).toBe('teams');
  });

  test('"channel meeting" → channel', () => {
    expect(parseMeetingType('channel meeting')).toBe('channel');
  });

  test('"Teams channel" → channel', () => {
    expect(parseMeetingType('Teams channel discussion')).toBe('channel');
  });

  test('"external meeting" → external', () => {
    expect(parseMeetingType('external meeting with the vendor')).toBe('external');
  });

  test('"client meeting" → external', () => {
    expect(parseMeetingType('client meeting')).toBe('external');
  });

  test('"vendor call" → external', () => {
    expect(parseMeetingType('vendor call')).toBe('external');
  });

  test('"outside the company" → external', () => {
    expect(parseMeetingType('meeting outside the company')).toBe('external');
  });

  test('unrecognised input → unknown', () => {
    expect(parseMeetingType('just a regular catch-up')).toBe('unknown');
  });

  test('empty string → unknown', () => {
    expect(parseMeetingType('')).toBe('unknown');
  });
});
