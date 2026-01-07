// Central pause/resume flag for frontend/backend sync automations.

namespace PauseService {
  const KEY = Config.PROPERTY_KEYS.FRONTEND_SYNC_PAUSED;

  function props() {
    return Config.scriptProperties();
  }

  export function isPaused(): boolean {
    try {
      const raw = props().getProperty(KEY) || '';
      return String(raw).toLowerCase() === 'true';
    } catch (err) {
      Log.warn(`Unable to read pause flag: ${err}`);
      return false;
    }
  }

  export function pause(reason?: string) {
    try {
      const payload = reason ? JSON.stringify({ paused: true, reason, at: new Date().toISOString() }) : 'true';
      props().setProperty(KEY, payload);
    } catch (err) {
      Log.warn(`Unable to set pause flag: ${err}`);
    }
  }

  export function resume(): boolean {
    const wasPaused = isPaused();
    try {
      props().deleteProperty(KEY);
    } catch (err) {
      Log.warn(`Unable to clear pause flag: ${err}`);
    }
    return wasPaused;
  }

  export function pauseInfo(): string {
    try {
      const raw = props().getProperty(KEY) || '';
      if (!raw) return 'not paused';
      if (raw === 'true') return 'paused';
      try {
        const parsed = JSON.parse(raw);
        if (parsed?.reason) return `paused: ${parsed.reason}`;
      } catch {
        // fall through
      }
      return 'paused';
    } catch (err) {
      Log.warn(`Unable to read pause info: ${err}`);
      return 'unknown';
    }
  }
}
