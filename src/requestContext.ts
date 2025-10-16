//src/requestContext.ts
import { AsyncLocalStorage } from "node:async_hooks";
import type { IncomingHttpHeaders } from "node:http";

export interface RequestContext {
  headers: IncomingHttpHeaders;
}

export const requestContext = new AsyncLocalStorage<RequestContext>();

export function getHeader(name: string): string | undefined {
  const store = requestContext.getStore();
  const v = store?.headers?.[name.toLowerCase()];
  return Array.isArray(v) ? v[0] : v;
}

export function requireHeader(name: string): string {
  const value = getHeader(name);
  if (!value) {
    throw new Error(`Missing required header: ${name}`);
  }
  return value;
}
