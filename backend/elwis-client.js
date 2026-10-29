// Schlanker Elwis-Client für das Add-in-Backend.
// SICHERHEIT: Stufe 1 ist strikt lesend. Erlaubt sind nur GET-Aufrufe und die
// reinen Abfrage-Endpunkte (POST .../list). Alles andere wird hier blockiert.

const READ_ONLY_POST = new Set(['/api/tenant/artifact/list', '/api/tenant/link/list', '/api/auth/login']);

export class ElwisClient {
  constructor(baseUrl, cookie = null) {
    this.baseUrl = baseUrl.replace(/\/$/, '');
    this.cookie = cookie;
  }

  async login(email, password) {
    const res = await fetch(this.baseUrl + '/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password }),
      redirect: 'manual'
    });
    if (res.status === 403) throw httpError(403, 'Konto gesperrt');
    if (!res.ok) throw httpError(401, 'Anmeldung fehlgeschlagen');
    const setCookie = res.headers.getSetCookie ? res.headers.getSetCookie() : [res.headers.get('set-cookie')].filter(Boolean);
    this.cookie = setCookie.map((c) => c.split(';')[0]).join('; ');
    if (!this.cookie) throw httpError(500, 'Keine Sitzung erhalten');
    return this.get('/api/tenant/user/info');
  }

  async request(method, path, body) {
    if (method !== 'GET' && !READ_ONLY_POST.has(path)) {
      throw httpError(405, `Schreibzugriff blockiert: ${method} ${path}`);
    }
    const res = await fetch(this.baseUrl + path, {
      method,
      headers: { 'Content-Type': 'application/json', Cookie: this.cookie || '' },
      body: body ? JSON.stringify(body) : undefined
    });
    if (res.status === 401) throw httpError(401, 'Elwis-Sitzung abgelaufen');
    if (!res.ok) throw httpError(res.status, `Elwis ${res.status} bei ${path}`);
    return res.json();
  }

  get(path) { return this.request('GET', path); }

  // Mongo-artige Filter, wie sie die Elwis-Oberfläche verwendet
  async listArtifacts(filter, { limit = 200, skip = 0, sort = { 'created.on': -1 } } = {}) {
    const j = await this.request('POST', '/api/tenant/artifact/list', {
      limit: String(limit), skip: String(skip), filter: JSON.stringify(filter), sort: JSON.stringify(sort)
    });
    return j.data || [];
  }

  async listLinks(filter, { limit = 500 } = {}) {
    const j = await this.request('POST', '/api/tenant/link/list', { limit: String(limit), filter: JSON.stringify(filter) });
    return j.data || [];
  }
}

export const oid = (id) => ({ $oid: id });
export const oids = (ids) => ({ $in: ids.map(oid) });

export function httpError(status, message) {
  return Object.assign(new Error(message), { status });
}
