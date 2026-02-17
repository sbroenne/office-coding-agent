function extractRetryAfterSeconds(message: string): number | null {
  const match = /retry\s+after\s+(\d+)\s*seconds?/i.exec(message);
  if (!match) return null;

  const parsed = Number.parseInt(match[1] ?? '', 10);
  return Number.isFinite(parsed) ? parsed : null;
}

export function normalizeChatErrorMessage(message: string): string {
  const trimmed = message.trim();
  const retryAfterSeconds = extractRetryAfterSeconds(trimmed);

  if (/\b429\b|too many requests|rate limit|token rate limit/i.test(trimmed)) {
    const retry = retryAfterSeconds ? ` Retry in about ${retryAfterSeconds}s.` : '';
    return (
      'Rate limit reached for the current model/deployment. ' +
      'Try a smaller request, switch models, or wait briefly.' +
      retry
    );
  }

  if (/\b404\b|not found|deployment.*not found|resource.*not found/i.test(trimmed)) {
    return (
      'Endpoint or deployment not found. Check your endpoint URL and model deployment name in Settings.'
    );
  }

  if (/\b401\b|unauthorized|invalid api key|authentication/i.test(trimmed)) {
    return 'Authentication failed. Verify your API key in Settings.';
  }

  if (/\b403\b|forbidden|permission|access denied/i.test(trimmed)) {
    return 'Access denied for this endpoint/model. Verify permissions and deployment access.';
  }

  return trimmed;
}
