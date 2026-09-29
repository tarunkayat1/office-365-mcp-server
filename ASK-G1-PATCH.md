# Ask G1 token bridge patch

Pinned commit: `12c9dd358e00576d6b2f18a9480891050a23851c`

## Changes

- `auth/token-manager.js`: request-scoped Bearer token via `AsyncLocalStorage` and `runWithAccessToken()`
- `index.js`: extract `Authorization: Bearer` from HTTP/SSE requests and propagate per session

When Ask G1 injects a user's Microsoft Graph access token in the `Authorization` header,
the sidecar uses it instead of `~/.office-mcp-tokens.json`. File-based tokens remain
available for standalone local development when no header is present.
