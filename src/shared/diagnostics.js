const STORAGE_KEY = "agendaEnforcerDiagnostics";
const MAX_ENTRIES = 200;

function getStorage() {
  if (typeof window === "undefined" || !window.localStorage) {
    return null;
  }

  return window.localStorage;
}

function normalizeDetails(details) {
  if (details === undefined || details === null) {
    return "";
  }

  if (typeof details === "string") {
    return details;
  }

  try {
    return JSON.stringify(details);
  } catch (error) {
    return String(details);
  }
}

export function readDiagnostics() {
  const storage = getStorage();

  if (!storage) {
    return [];
  }

  try {
    const raw = storage.getItem(STORAGE_KEY);
    if (!raw) {
      return [];
    }

    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

export function writeDiagnostics(entries) {
  const storage = getStorage();

  if (!storage) {
    return;
  }

  try {
    storage.setItem(STORAGE_KEY, JSON.stringify(entries.slice(-MAX_ENTRIES)));
  } catch (error) {
    // Ignore storage failures so diagnostics never block add-in execution.
  }
}

export function addDiagnostic(stage, details) {
  const entries = readDiagnostics();

  entries.push({
    timestamp: new Date().toISOString(),
    stage: stage,
    details: normalizeDetails(details),
    location: typeof window !== "undefined" ? window.location.href : ""
  });

  writeDiagnostics(entries);
}

export function clearDiagnostics() {
  writeDiagnostics([]);
}
