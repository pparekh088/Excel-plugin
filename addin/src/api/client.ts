/**
 * Typed client for the Ledger Agent API tool-call envelope.
 */

import type { AuthProvider } from "../auth/provider";

export interface Health {
  status: "ok";
  version: string;
  auth_mode: string;
  session_store: string;
  tools: string[];
}

export interface SessionCreated {
  session_id: string;
  created_at: string;
  auth_mode: string;
}

export interface ToolCallAccepted {
  tool_call_id: string;
  tool: string;
  params: Record<string, unknown>;
  access: "read" | "write" | "control";
  status: "validated" | "completed" | "failed";
}

export interface ToolResultAck {
  tool_call_id: string;
  status: "validated" | "completed" | "failed";
  summary: string;
}

export interface ToolErrorBody {
  code: string;
  message: string;
  detail?: Record<string, unknown> | null;
}

export class ApiError extends Error {
  constructor(
    public readonly status: number,
    public readonly body: unknown,
    message?: string
  ) {
    super(message ?? `API error ${status}`);
    this.name = "ApiError";
  }
}

export class LedgerClient {
  constructor(
    private readonly baseUrl: string,
    private readonly auth: AuthProvider
  ) {}

  private async request<T>(method: string, path: string, body?: unknown): Promise<T> {
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    const token = await this.auth.getAccessToken();
    if (token) headers["Authorization"] = `Bearer ${token}`;

    const response = await fetch(`${this.baseUrl}${path}`, {
      method,
      headers,
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    const text = await response.text();
    const json: unknown = text ? JSON.parse(text) : null;
    if (!response.ok) {
      throw new ApiError(response.status, json);
    }
    return json as T;
  }

  health(): Promise<Health> {
    return this.request("GET", "/healthz");
  }

  createSession(): Promise<SessionCreated> {
    return this.request("POST", "/api/v1/sessions");
  }

  createToolCall(
    sessionId: string,
    tool: string,
    params: Record<string, unknown>
  ): Promise<ToolCallAccepted> {
    return this.request("POST", `/api/v1/sessions/${sessionId}/tool-calls`, { tool, params });
  }

  postToolResult(
    sessionId: string,
    toolCallId: string,
    body:
      | { ok: true; result: Record<string, unknown> }
      | { ok: false; error: ToolErrorBody }
  ): Promise<ToolResultAck> {
    return this.request(
      "POST",
      `/api/v1/sessions/${sessionId}/tool-calls/${toolCallId}/result`,
      body
    );
  }
}
