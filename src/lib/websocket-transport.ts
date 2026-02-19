/*---------------------------------------------------------------------------------------------
 *  WebSocket transport for browser environments
 *  Custom transport layer for vscode-jsonrpc over WebSocket
 *  Source: https://github.com/patniko/github-copilot-office
 *--------------------------------------------------------------------------------------------*/

import {
  AbstractMessageReader,
  AbstractMessageWriter,
  type DataCallback,
  type Disposable,
  type Message,
  type MessageReader,
  type MessageWriter,
} from 'vscode-jsonrpc';

export class WebSocketMessageReader extends AbstractMessageReader implements MessageReader {
  private buffer = new Uint8Array(0);
  private callback: DataCallback | null = null;
  private decoder = new TextDecoder();

  constructor(private socket: WebSocket) {
    super();
    this.socket.addEventListener('message', event => {
      void (async () => {
        let bytes: Uint8Array;
        if (event.data instanceof Blob) {
          bytes = new Uint8Array(await event.data.arrayBuffer());
        } else if (event.data instanceof ArrayBuffer) {
          bytes = new Uint8Array(event.data);
        } else {
          bytes = new TextEncoder().encode(event.data as string);
        }

        const newBuffer = new Uint8Array(this.buffer.length + bytes.length);
        newBuffer.set(this.buffer);
        newBuffer.set(bytes, this.buffer.length);
        this.buffer = newBuffer;

        this.processBuffer();
      })();
    });

    this.socket.addEventListener('error', () => {
      this.fireError(new Error('WebSocket error'));
    });

    this.socket.addEventListener('close', () => {
      this.fireClose();
    });
  }

  private processBuffer(): void {
    while (true) {
      const headerEndIndex = this.findSequence(this.buffer, [13, 10, 13, 10]); // \r\n\r\n
      if (headerEndIndex === -1) break;

      const header = this.decoder.decode(this.buffer.slice(0, headerEndIndex));
      const match = /Content-Length:\s*(\d+)/i.exec(header);
      if (!match) break;

      const contentLength = parseInt(match[1], 10);
      const contentStart = headerEndIndex + 4;
      const messageEnd = contentStart + contentLength;

      if (this.buffer.length < messageEnd) break;

      const content = this.decoder.decode(this.buffer.slice(contentStart, messageEnd));
      this.buffer = this.buffer.slice(messageEnd);

      try {
        const msg = JSON.parse(content) as Message;
        this.callback?.(msg);
      } catch {
        // Skip malformed JSON
      }
    }
  }

  private findSequence(arr: Uint8Array, seq: number[]): number {
    outer: for (let i = 0; i <= arr.length - seq.length; i++) {
      for (let j = 0; j < seq.length; j++) {
        if (arr[i + j] !== seq[j]) continue outer;
      }
      return i;
    }
    return -1;
  }

  listen(callback: DataCallback): Disposable {
    this.callback = callback;
    return {
      dispose: () => {
        this.callback = null;
      },
    };
  }
}

export class WebSocketMessageWriter extends AbstractMessageWriter implements MessageWriter {
  private errorCount = 0;

  constructor(private socket: WebSocket) {
    super();
  }

  write(msg: Message): Promise<void> {
    try {
      const content = JSON.stringify(msg);
      const header = `Content-Length: ${new TextEncoder().encode(content).length}\r\n\r\n`;
      this.socket.send(header + content);
    } catch (error) {
      this.errorCount++;
      this.fireError(
        error instanceof Error ? error : new Error(String(error)),
        msg,
        this.errorCount
      );
    }
    return Promise.resolve();
  }

  end(): void {
    // WebSocket close is handled externally
  }
}
