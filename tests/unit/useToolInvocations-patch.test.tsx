/**
 * Unit test for @assistant-ui/react useToolInvocations argsText handling.
 *
 * Validates that when streaming tool call argsText transitions from
 * incomplete JSON (keys in LLM generation order) to complete JSON
 * (keys reordered by JSON.stringify), or to entirely different text,
 * the hook does NOT throw — it gracefully restarts the args stream.
 */
import { describe, it, expect, vi } from 'vitest';
import React, { useState } from 'react';
import { render, act } from '@testing-library/react';
import { INTERNAL } from '@assistant-ui/react';

const { useToolInvocations } = INTERNAL;

/**
 * Build a minimal ThreadMessage-like object for the hook.
 * The hook only reads `content[].type`, `.toolCallId`, `.toolName`,
 * `.argsText`, and `.result`.
 */
function makeState(
  toolCalls: Array<{
    toolCallId: string;
    toolName: string;
    argsText: string;
    result?: unknown;
  }>,
  isRunning = true
) {
  return {
    messages: [
      {
        role: 'assistant' as const,
        id: 'msg-1',
        createdAt: new Date(),
        status: { type: 'running' as const },
        content: toolCalls.map(tc => ({
          type: 'tool-call' as const,
          ...tc,
        })),
        metadata: {},
      },
    ],
    isRunning,
  };
}

/**
 * Wrapper component that drives useToolInvocations with externally
 * controlled state, so we can simulate streaming → complete transitions.
 */
function HookDriver({
  stateRef,
  onRender,
}: {
  stateRef: React.MutableRefObject<ReturnType<typeof makeState>>;
  onRender?: () => void;
}) {
  // Force re-render by bumping a counter
  const [, setTick] = useState(0);

  // Expose a way to trigger re-render
  (stateRef as any).rerender = () => setTick(t => t + 1);

  useToolInvocations({
    state: stateRef.current as any,
    getTools: () => undefined,
    onResult: vi.fn(),
    setToolStatuses: vi.fn(),
  });

  onRender?.();
  return null;
}

describe('useToolInvocations patch', () => {
  it('does not throw when argsText transitions from incomplete to complete JSON with reordered keys', () => {
    const stateRef = { current: makeState([]) } as any;

    // 1. Initial render with empty messages (sets isInitialState to false)
    const { rerender } = render(<HookDriver stateRef={stateRef} />);

    // 2. Simulate streaming: incomplete argsText with keys in LLM order
    //    Note: sheetName comes first during streaming
    const streamingArgsText =
      '{"sheetName":"Sheet1","address":"A1:C1","formulas":[["Name","Age","City"';

    act(() => {
      stateRef.current = makeState([
        {
          toolCallId: 'tc-1',
          toolName: 'setRange',
          argsText: streamingArgsText,
        },
      ]);
      rerender(<HookDriver stateRef={stateRef} />);
    });

    // 3. Simulate completion: complete argsText with REORDERED keys
    //    JSON.stringify puts "address" first, then "formulas", then "sheetName"
    const completeArgsText =
      '{"address":"A1:C1","formulas":[["Name","Age","City"]],"sheetName":"Sheet1"}';

    // Without the patch, this would throw:
    //   "Tool call argsText can only be appended, not updated"
    expect(() => {
      act(() => {
        stateRef.current = makeState([
          {
            toolCallId: 'tc-1',
            toolName: 'setRange',
            argsText: completeArgsText,
          },
        ]);
        rerender(<HookDriver stateRef={stateRef} />);
      });
    }).not.toThrow();
  });

  it('still works for normal append-only streaming (no key reorder)', () => {
    const stateRef = { current: makeState([]) } as any;

    const { rerender } = render(<HookDriver stateRef={stateRef} />);

    // Stream chunk 1
    act(() => {
      stateRef.current = makeState([
        {
          toolCallId: 'tc-2',
          toolName: 'setRange',
          argsText: '{"address":"A1"',
        },
      ]);
      rerender(<HookDriver stateRef={stateRef} />);
    });

    // Stream chunk 2 (appends to chunk 1)
    act(() => {
      stateRef.current = makeState([
        {
          toolCallId: 'tc-2',
          toolName: 'setRange',
          argsText: '{"address":"A1","value":"hello"}',
        },
      ]);
      rerender(<HookDriver stateRef={stateRef} />);
    });

    // No error expected
  });

  it('does not throw for non-appendable argsText (handled via replacement stream)', () => {
    const stateRef = { current: makeState([]) } as any;

    const { rerender } = render(<HookDriver stateRef={stateRef} />);

    // Stream chunk 1
    act(() => {
      stateRef.current = makeState([
        {
          toolCallId: 'tc-3',
          toolName: 'setRange',
          argsText: '{"address":"A1"',
        },
      ]);
      rerender(<HookDriver stateRef={stateRef} />);
    });

    // Non-appendable chunk — handled gracefully via replacement stream
    expect(() => {
      act(() => {
        stateRef.current = makeState([
          {
            toolCallId: 'tc-3',
            toolName: 'setRange',
            argsText: 'CORRUPTED_GARBAGE',
          },
        ]);
        rerender(<HookDriver stateRef={stateRef} />);
      });
    }).not.toThrow();
  });
});
