export class SessionCache {

  public static set(siteId: string, value: string, expireTime: number = 5 * 60 * 1000): void {
    const record = {
      value: value,
      expiresAt: Date.now() + expireTime
    };
    localStorage.setItem(siteId, JSON.stringify(record));
  }

  public static get(siteId: string): string | undefined {
    const item = localStorage.getItem(siteId);
    if (!item) return undefined;

    try {
      const record = JSON.parse(item);
      if (Date.now() > record.expiresAt) {
        localStorage.removeItem(siteId);
        return undefined;
      }
      return record.value;
    } catch {
      localStorage.removeItem(siteId);
      return undefined;
    }
  }
}