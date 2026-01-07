// Header utilities: machine header (row 1) -> display header (row 2).

namespace Headers {
  const ACRONYMS = new Set(['id', 'as', 'dob', 'cip', 'afsc', 'llab', 'pt', 'pdf', 'mfr', 'n/a']);

  const WORD_OVERRIDES: { [key: string]: string } = {
    // Common fields
    last: 'Last',
    first: 'First',
    name: 'Name',
    email: 'Email',
    phone: 'Phone',
    cell: 'Cell',
    office: 'Office',
    location: 'Location',
    town: 'Town',
    state: 'State',
    year: 'Year',
    term: 'Term',
    status: 'Status',
    notes: 'Notes',

    // Domain-specific
    dob: 'DOB',
    cip: 'CIP',
    afsc: 'AFSC',
    llab: 'LLAB',
    as: 'AS',
    id: 'ID',
    datetime: 'Date/Time',
    pct: '%',

    // Composite hints
    submitted: 'Submitted',
    decided: 'Decided',
    attendance: 'Attendance',
    excusal: 'Excusal',
    requested: 'Requested',
    denied: 'Denied',
    approved: 'Approved',
  };

  function capitalizeWord(word: string): string {
    const lower = word.toLowerCase();
    if (WORD_OVERRIDES[lower]) return WORD_OVERRIDES[lower];
    if (ACRONYMS.has(lower)) return lower.toUpperCase();
    if (lower.length === 0) return '';
    return lower.charAt(0).toUpperCase() + lower.slice(1);
  }

  export function humanizeHeader(machineHeader: string): string {
    const raw = (machineHeader || '').trim();
    if (!raw) return '';

    // Special case: keep N/A as-is.
    if (raw.toLowerCase() === 'n/a') return 'N/A';

    const parts = raw
      .replace(/\s+/g, '_')
      .split('_')
      .filter((p) => p.length > 0);

    const words = parts.map(capitalizeWord);

    // Clean up cases like "Attendance %" where we want "%" to attach.
    const joined = words.join(' ');
    return joined
      .replace(/\s+%/g, ' %')
      .replace(/Date\/Time/g, 'Date/Time');
  }

  export function humanizeHeaders(machineHeaders: string[]): string[] {
    return machineHeaders.map((h) => humanizeHeader(h));
  }
}
