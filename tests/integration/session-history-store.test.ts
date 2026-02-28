/**
 * Integration tests for sessionHistoryStore.
 *
 * Validates createSession, upsertActiveSession, deleteSession, clearSessionsForHost,
 * 50-item limit (MAX_SESSIONS), host filtering, and active session management.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { useSessionHistoryStore } from '@/stores/sessionHistoryStore';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';

beforeEach(() => {
  useSessionHistoryStore.setState({
    sessions: [],
    activeSessionId: null,
  });
});

// ─── createSession ────────────────────────────────────────────────────────────

describe('sessionHistoryStore — createSession', () => {
  it('creates a new session and sets it as active', () => {
    const id = useSessionHistoryStore.getState().createSession('excel');
    expect(typeof id).toBe('string');
    expect(id.length).toBeGreaterThan(0);
    expect(useSessionHistoryStore.getState().activeSessionId).toBe(id);
    expect(useSessionHistoryStore.getState().sessions).toHaveLength(1);
  });

  it('new session has correct defaults', () => {
    const id = useSessionHistoryStore.getState().createSession('powerpoint');
    const session = useSessionHistoryStore.getState().sessions.find(s => s.id === id);
    expect(session).toBeDefined();
    expect(session!.title).toBe('New conversation');
    expect(session!.host).toBe('powerpoint');
    expect(session!.messages).toEqual([]);
    expect(typeof session!.updatedAt).toBe('number');
  });

  it('creating multiple sessions accumulates them', () => {
    useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().createSession('powerpoint');
    expect(useSessionHistoryStore.getState().sessions).toHaveLength(3);
  });
});

// ─── setActiveSession ─────────────────────────────────────────────────────────

describe('sessionHistoryStore — setActiveSession', () => {
  it('sets the active session id', () => {
    const id1 = useSessionHistoryStore.getState().createSession('excel');
    const id2 = useSessionHistoryStore.getState().createSession('excel');
    expect(useSessionHistoryStore.getState().activeSessionId).toBe(id2);

    useSessionHistoryStore.getState().setActiveSession(id1);
    expect(useSessionHistoryStore.getState().activeSessionId).toBe(id1);
  });
});

// ─── upsertActiveSession ─────────────────────────────────────────────────────

describe('sessionHistoryStore — upsertActiveSession', () => {
  it('creates a session if none is active, then upserts it', () => {
    useSessionHistoryStore.getState().upsertActiveSession({
      host: 'excel',
      title: 'My Chat',
      messages: [{ role: 'user', content: 'hello' }],
    });

    const { sessions, activeSessionId } = useSessionHistoryStore.getState();
    expect(sessions).toHaveLength(1);
    expect(activeSessionId).toBe(sessions[0].id);
    expect(sessions[0].title).toBe('My Chat');
    expect(sessions[0].messages).toHaveLength(1);
  });

  it('updates existing active session on subsequent calls', () => {
    const id = useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().upsertActiveSession({
      host: 'excel',
      title: 'Updated Title',
      messages: [{ role: 'assistant', content: 'hi' }],
    });

    const session = useSessionHistoryStore.getState().sessions.find(s => s.id === id);
    expect(session!.title).toBe('Updated Title');
    expect(session!.messages).toHaveLength(1);
  });

  it('upsert moves session to the top (most recent)', () => {
    const id1 = useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().createSession('excel');

    // Switch back to id1 and upsert
    useSessionHistoryStore.getState().setActiveSession(id1);
    useSessionHistoryStore.getState().upsertActiveSession({
      host: 'excel',
      title: 'Refreshed',
      messages: [],
    });

    const sessions = useSessionHistoryStore.getState().sessions;
    // After sorting by updatedAt desc, id1 should be first
    expect(sessions[0].id).toBe(id1);
  });

  it('deep-clones messages to prevent mutation (toSerializableMessages)', () => {
    const original = [{ role: 'user', content: 'hello', nested: { a: 1 } }];
    useSessionHistoryStore.getState().upsertActiveSession({
      host: 'excel',
      title: 'Clone Test',
      messages: original,
    });

    const stored = useSessionHistoryStore.getState().sessions[0].messages as typeof original;
    expect(stored).toEqual(original);
    // Should be a different reference (deep clone)
    expect(stored).not.toBe(original);
  });
});

// ─── deleteSession ────────────────────────────────────────────────────────────

describe('sessionHistoryStore — deleteSession', () => {
  it('removes the specified session', () => {
    const id = useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().deleteSession(id);
    expect(useSessionHistoryStore.getState().sessions).toHaveLength(0);
  });

  it('clears activeSessionId when the active session is deleted (falls back to first remaining)', () => {
    const id1 = useSessionHistoryStore.getState().createSession('excel');
    const id2 = useSessionHistoryStore.getState().createSession('excel');

    // id2 is active
    useSessionHistoryStore.getState().deleteSession(id2);
    // Should fall back to id1
    expect(useSessionHistoryStore.getState().activeSessionId).toBe(id1);
  });

  it('sets activeSessionId to null when last session is deleted', () => {
    const id = useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().deleteSession(id);
    expect(useSessionHistoryStore.getState().activeSessionId).toBeNull();
  });

  it('preserves activeSessionId when a non-active session is deleted', () => {
    const id1 = useSessionHistoryStore.getState().createSession('excel');
    const id2 = useSessionHistoryStore.getState().createSession('excel');
    // id2 is active
    useSessionHistoryStore.getState().deleteSession(id1);
    expect(useSessionHistoryStore.getState().activeSessionId).toBe(id2);
  });
});

// ─── clearSessionsForHost ─────────────────────────────────────────────────────

describe('sessionHistoryStore — clearSessionsForHost', () => {
  it('removes all sessions for the specified host', () => {
    useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().createSession('powerpoint');

    useSessionHistoryStore.getState().clearSessionsForHost('excel');

    const sessions = useSessionHistoryStore.getState().sessions;
    expect(sessions).toHaveLength(1);
    expect(sessions[0].host).toBe('powerpoint');
  });

  it('updates activeSessionId when cleared host contained active session', () => {
    useSessionHistoryStore.getState().createSession('powerpoint');
    useSessionHistoryStore.getState().createSession('excel');
    // excel session is active

    useSessionHistoryStore.getState().clearSessionsForHost('excel');
    const { activeSessionId, sessions } = useSessionHistoryStore.getState();
    // Should fall back to the powerpoint session
    expect(sessions).toHaveLength(1);
    expect(activeSessionId).toBe(sessions[0].id);
  });

  it('preserves activeSessionId when another host is cleared', () => {
    useSessionHistoryStore.getState().createSession('excel');
    const excelId = useSessionHistoryStore.getState().activeSessionId;
    useSessionHistoryStore.getState().createSession('powerpoint');
    useSessionHistoryStore.getState().setActiveSession(excelId!);

    useSessionHistoryStore.getState().clearSessionsForHost('powerpoint');
    expect(useSessionHistoryStore.getState().activeSessionId).toBe(excelId);
  });

  it('sets activeSessionId to null when all sessions cleared', () => {
    useSessionHistoryStore.getState().createSession('excel');
    useSessionHistoryStore.getState().clearSessionsForHost('excel');
    expect(useSessionHistoryStore.getState().activeSessionId).toBeNull();
  });
});

// ─── MAX_SESSIONS limit ──────────────────────────────────────────────────────

describe('sessionHistoryStore — MAX_SESSIONS (50) limit', () => {
  it('trims sessions to 50 when limit is exceeded', () => {
    // Create 55 sessions
    for (let i = 0; i < 55; i++) {
      useSessionHistoryStore.getState().createSession('excel');
    }

    const sessions = useSessionHistoryStore.getState().sessions;
    expect(sessions.length).toBeLessThanOrEqual(50);
  });

  it('keeps the most recent sessions when trimming', () => {
    // Seed with known timestamps
    const sessions: SessionHistoryItem[] = [];
    for (let i = 0; i < 55; i++) {
      sessions.push({
        id: `session-${i}`,
        title: `Session ${i}`,
        host: 'excel',
        updatedAt: 1000 + i, // older = lower number
        messages: [],
      });
    }
    useSessionHistoryStore.setState({ sessions, activeSessionId: 'session-54' });

    // Trigger trimming via upsert
    useSessionHistoryStore.getState().upsertActiveSession({
      host: 'excel',
      title: 'Newest',
      messages: [],
    });

    const stored = useSessionHistoryStore.getState().sessions;
    expect(stored.length).toBeLessThanOrEqual(50);
    // The first (session-0) through (session-4) should have been trimmed
    const ids = stored.map(s => s.id);
    expect(ids).not.toContain('session-0');
    expect(ids).not.toContain('session-1');
  });
});

// ─── Persistence configuration ───────────────────────────────────────────────

describe('sessionHistoryStore — persistence', () => {
  it('uses the correct persist key', () => {
    // Access the persist API to check the name
    const persistOptions = (
      useSessionHistoryStore as unknown as { persist: { getOptions: () => { name: string } } }
    ).persist;
    if (persistOptions?.getOptions) {
      expect(persistOptions.getOptions().name).toBe('office-coding-agent-session-history');
    }
  });
});
