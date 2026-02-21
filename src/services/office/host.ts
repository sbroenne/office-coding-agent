export type OfficeHostApp = 'excel' | 'powerpoint' | 'word' | 'unknown';

function normalizeHost(value: string | undefined): OfficeHostApp {
  const host = value?.toLowerCase();
  if (host === 'excel') return 'excel';
  if (host === 'powerpoint') return 'powerpoint';
  if (host === 'word') return 'word';
  return 'unknown';
}

export function detectOfficeHost(): OfficeHostApp {
  if (typeof Office === 'undefined') {
    return 'excel';
  }

  const hostValue = Office.context?.host;
  if (hostValue == null) {
    return 'excel';
  }

  if (typeof hostValue === 'string') {
    return normalizeHost(hostValue);
  }

  const hostType = Office.HostType;
  if (!hostType) {
    return 'excel';
  }

  if (hostValue === hostType.Excel) return 'excel';
  if (hostValue === hostType.PowerPoint) return 'powerpoint';
  if (hostValue === hostType.Word) return 'word';
  return 'unknown';
}
