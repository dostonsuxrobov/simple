/**
 * The name Simple Docs puts on comments and suggestions: the Windows user name
 * until the person chooses another one once (stored on this computer only).
 *
 * No DOM or runtime imports: Node tests load this module directly.
 */

export const AUTHOR_STORAGE_KEY = "simple-docs:author-name";
export const FALLBACK_AUTHOR_NAME = "Author";
const MAX_NAME_LENGTH = 80;

export interface AuthorIdentity {
  id: string;
  firstName: string;
  lastName: string;
}

interface StorageLike {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

/** A tidy display name: trimmed, single spaces, no control characters, at most 80 characters. */
export function cleanAuthorName(raw: unknown): string {
  return String(raw ?? "")
    .replace(/[\u0000-\u001f\u007f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, MAX_NAME_LENGTH)
    .trim();
}

/** "DOMAIN\\name" or "name@domain" (as Windows can report them) shown as "name". */
export function windowsUserDisplayName(raw: unknown): string {
  let name = cleanAuthorName(raw);
  if (name.includes("\\")) name = name.slice(name.lastIndexOf("\\") + 1);
  if (/^[^\s@]+@[^\s@]+$/.test(name)) name = name.slice(0, name.indexOf("@"));
  return cleanAuthorName(name);
}

const SHARED_PROFILES = new Set(["public", "default", "default user", "all users", "defaultuser0", "wdagutilityaccount"]);

/**
 * The Windows user name from the app's own location when it runs from a user
 * profile (C:\Users\<name>\…, which is where the portable app unpacks), for
 * when main does not report the name itself. Shared profiles never count.
 */
export function windowsUserFromAppUrl(url: string): string | null {
  const match = /^file:\/\/\/[a-z]:\/users\/([^/]+)\//i.exec(String(url ?? ""));
  if (!match) return null;
  let name: string;
  try {
    name = decodeURIComponent(match[1]);
  } catch {
    return null;
  }
  const clean = cleanAuthorName(name);
  return clean && !SHARED_PROFILES.has(clean.toLowerCase()) ? clean : null;
}

/** WordCanvas identity fields; the engine shows initials from the first letters of both names. */
export function authorIdentity(name: string, id = "local-user"): AuthorIdentity {
  const clean = cleanAuthorName(name) || FALLBACK_AUTHOR_NAME;
  const space = clean.indexOf(" ");
  return space < 0
    ? { id, firstName: clean, lastName: "" }
    : { id, firstName: clean.slice(0, space), lastName: clean.slice(space + 1) };
}

export function authorDisplayName(identity: Pick<AuthorIdentity, "firstName" | "lastName">): string {
  return cleanAuthorName(`${identity.firstName ?? ""} ${identity.lastName ?? ""}`);
}

/** The name the person chose, or null when they never chose one. */
export function readStoredAuthorName(storage: StorageLike | null | undefined): string | null {
  try {
    const stored = cleanAuthorName(storage?.getItem(AUTHOR_STORAGE_KEY) ?? "");
    return stored || null;
  } catch {
    return null;
  }
}

/** Remembers the chosen name; an empty name forgets it. Returns false when storage is unavailable. */
export function storeAuthorName(storage: StorageLike | null | undefined, name: string): boolean {
  try {
    const clean = cleanAuthorName(name);
    if (clean) storage?.setItem(AUTHOR_STORAGE_KEY, clean);
    else storage?.removeItem(AUTHOR_STORAGE_KEY);
    return Boolean(storage);
  } catch {
    return false;
  }
}

/** The name to use: the chosen one, else the Windows user name, else a neutral fallback. */
export function resolveAuthorName(stored: string | null | undefined, windowsUser: string | null | undefined): { name: string; chosen: boolean } {
  const chosen = cleanAuthorName(stored);
  if (chosen) return { name: chosen, chosen: true };
  const system = windowsUserDisplayName(windowsUser);
  return { name: system || FALLBACK_AUTHOR_NAME, chosen: false };
}
