import React from 'react';

interface Props {
  children: React.ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

/**
 * Error boundary around the chat UI.
 * If a render error occurs in Thread/MarkdownText/ToolFallback,
 * this keeps the header and settings functional.
 */
export class ChatErrorBoundary extends React.Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  override componentDidCatch(error: Error, info: React.ErrorInfo) {
    console.error('[ChatErrorBoundary] Render error:', error, info.componentStack);
  }

  override render() {
    if (this.state.hasError) {
      const errorMessage = this.state.error?.message?.trim();

      return (
        <div className="flex flex-1 flex-col items-center justify-center gap-3 p-4 text-center">
          <p className="text-sm font-medium text-destructive">Something went wrong</p>
          <p className="text-xs text-muted-foreground">
            {errorMessage && errorMessage.length > 0
              ? errorMessage
              : 'An unexpected error occurred.'}
          </p>
          <button
            onClick={() => this.setState({ hasError: false, error: null })}
            className="rounded-md bg-primary px-3 py-1.5 text-xs text-primary-foreground hover:bg-primary/90"
          >
            Try again
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
