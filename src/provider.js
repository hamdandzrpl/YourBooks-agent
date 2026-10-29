// Provider-agnostic chat-completions adapter.
//
// The agent speaks the OpenAI Chat Completions dialect because almost every
// provider exposes it: OpenAI, OpenRouter, DeepSeek, Groq, Together, Fireworks,
// Mistral, xAI, vikey.ai, Ollama/LM Studio, and the OpenAI-compatible endpoints
// of Anthropic, Gemini and Anthropic-compatible gateways. Providers still differ
// in small ways, so this module centralises the differences instead of letting
// them leak into the agent loop.

/** Providers whose auth header is not `Authorization: Bearer <key>`. */
const authStyles = [
  { match: /(^|\.)anthropic\.com$/i, header: "x-api-key", extra: { "anthropic-version": "2023-06-01" } },
  { match: /(^|\.)googleapis\.com$/i, header: "x-goog-api-key", extra: {} },
  { match: /(^|\.)azure\.com$/i, header: "api-key", extra: {} },
];

/** Providers that reject unknown fields such as `reasoning_effort`. */
function supportsReasoningEffort(url) {
  return /(^|\.)(xkiro\.com|openai\.com|openrouter\.ai|deepseek\.com|vikey\.ai|groq\.com|mistral\.ai|x\.ai|together\.xyz|fireworks\.ai)$/i
    .test(new URL(url).hostname);
}

export function providerHost(url) {
  try {
    return new URL(url).hostname;
  } catch {
    return "";
  }
}

export function buildHeaders(url, apiKey) {
  const headers = { "content-type": "application/json" };
  const host = providerHost(url);
  const style = authStyles.find((entry) => entry.match.test(host));
  if (style) {
    headers[style.header] = apiKey;
    Object.assign(headers, style.extra);
    // Anthropic-compatible gateways also accept the bearer form; sending both is
    // rejected, so the provider-specific header wins.
    return headers;
  }
  headers.authorization = `Bearer ${apiKey}`;
  return headers;
}

/**
 * Build the chat-completions request body for a provider.
 *
 * `tool_choice: "none"` is replaced by omitting tools entirely on providers that
 * do not understand the field, because a rejected body would fail the whole turn
 * at exactly the point where we are trying to wrap up gracefully.
 */
export function buildChatBody({ url, model, messages, tools, toolChoice, reasoningEffort, maxOutputTokens }) {
  const body = { model, messages };
  const canUseTools = Array.isArray(tools) && tools.length > 0;
  if (canUseTools && toolChoice !== "none") {
    body.tools = tools;
    body.tool_choice = toolChoice ?? "auto";
  } else if (canUseTools) {
    body.tools = tools;
    body.tool_choice = "none";
  }
  if (reasoningEffort && supportsReasoningEffort(url)) body.reasoning_effort = reasoningEffort;
  if (maxOutputTokens > 0) body.max_tokens = maxOutputTokens;
  return body;
}

/**
 * Normalise a provider response into `{ content, toolCalls }`.
 *
 * Covers the OpenAI shape plus the common deviations: `function_call` (legacy
 * single-call form), `tool_calls` with an object `arguments`, Gemini's
 * `functionCall` parts, and Anthropic-style content blocks.
 */
export function normalizeAssistantMessage(message) {
  if (!message || typeof message !== "object") return { content: "", toolCalls: [] };
  const content = typeof message.content === "string"
    ? message.content
    : Array.isArray(message.content)
      ? message.content.filter((part) => typeof part?.text === "string").map((part) => part.text).join("")
      : "";

  const toolCalls = [];
  const push = (name, args, id) => {
    if (!name) return;
    toolCalls.push({
      id: id || `call_${toolCalls.length}_${Math.random().toString(36).slice(2, 10)}`,
      type: "function",
      function: {
        name,
        arguments: typeof args === "string" ? args : JSON.stringify(args ?? {}),
      },
    });
  };

  for (const call of message.tool_calls ?? []) {
    push(call?.function?.name, call?.function?.arguments, call?.id);
  }
  if (message.function_call?.name) {
    push(message.function_call.name, message.function_call.arguments, message.function_call.id);
  }
  // Gemini / some gateways put calls inside content parts.
  if (Array.isArray(message.content)) {
    for (const part of message.content) {
      if (part?.functionCall?.name) push(part.functionCall.name, part.functionCall.args ?? {});
      if (part?.type === "tool_use" && part.name) push(part.name, part.input ?? {}, part.id);
    }
  }
  return { content, toolCalls };
}

/** Extract the assistant message from the many response envelopes in use. */
export function extractAssistantMessage(result) {
  return result?.choices?.[0]?.message
    ?? result?.choices?.[0]?.delta
    ?? result?.message
    ?? result?.response
    ?? null;
}

/** Human-readable provider error, including the body when the provider sends one. */
export function describeFailure(result, status) {
  const detail = result?.error?.message
    ?? result?.error?.type
    ?? result?.message
    ?? (typeof result === "string" ? result.slice(0, 300) : "");
  return `provider returned ${status}${detail ? `: ${detail}` : ""}`;
}
