# AGENTS.md — GitHub Copilot Office Add-in

## Project Overview

Microsoft Office Add-in integrating GitHub Copilot into Word, Excel, and PowerPoint.
Stack: React 19 + Fluent UI v9 (TypeScript frontend), Express.js (JS backend),
Electron (system tray), WebSocket + JSON-RPC (vscode-jsonrpc), Office.js API.

## Build / Dev / Test Commands

```bash
# Install dependencies
npm install

# Dev server (Express + Vite HMR, HTTPS on port 52390)
npm run dev

# Production build (Vite, outputs to dist/)
npm run build

# Production server (serves static dist/)
npm start

# Electron tray app
npm run start:tray

# Build desktop installers (current platform)
npm run build:installer
npm run build:installer:mac   # macOS .dmg
npm run build:installer:win   # Windows .exe

# Generate icon sizes
npm run build:icons
```

**No linter, formatter, or test framework is configured.** There are no eslint, prettier,
jest, vitest, or any automated test commands. Type-check with `npx tsc --noEmit`.

Node.js 20+ required (CI uses Node 22). Package manager: **npm** (lockfile: package-lock.json).

## Project Structure

```
src/
  server.js              # Dev server (Express + Vite HMR)
  server-prod.js         # Production server
  copilotProxy.js        # WebSocket <-> Copilot CLI (stdio) proxy
  tray/main.js           # Electron tray entry point
  ui/
    index.tsx            # Entry point (Office.onReady)
    App.tsx              # Main app component, state management hub
    components/          # React UI components (PascalCase files)
    lib/                 # Core client libraries (websocket, logging, permissions)
    tools/               # Office.js tool handlers (~31 tools)
      index.ts           # Tool registry: wordTools, powerpointTools, excelTools
assets/                  # Tray icons
certs/                   # SSL certs for localhost HTTPS
installer/               # macOS/Windows installer resources
scripts/                 # Build helper scripts
manifest.xml             # Office Add-in manifest
vite.config.js           # Vite build config (root: src/ui)
tsconfig.json            # TypeScript config (strict, ES2020, ESNext modules)
```

## Code Style Guidelines

### Formatting

- **2-space indentation**, no tabs
- **Semicolons always**
- **Double quotes** in TypeScript/TSX files; **single quotes** in JavaScript files
- **Trailing commas** in multiline objects, arrays, imports, and parameters
- Multi-line JSX attributes: one per line

### Imports

- All imports in a single block (no blank-line separation between groups)
- Implicit order: external libraries first, then local/relative imports
- **Named imports** exclusively — no default exports anywhere in the codebase
- Use `import type { ... }` for type-only imports
- Relative paths only (no path aliases like `@/` or `~/`)

```typescript
import { useState, useEffect, useRef, useCallback } from "react";
import { FluentProvider, makeStyles } from "@fluentui/react-components";
import type { Tool } from "@github/copilot-sdk";
import { ChatInput } from "./components/ChatInput";
import { remoteLog } from "./lib/remoteLog";
```

### Naming Conventions

| What | Convention | Example |
|------|-----------|---------|
| Variables, functions | camelCase | `handleSend`, `fetchModels` |
| React components | PascalCase | `ChatInput`, `MessageList` |
| Component files | PascalCase.tsx | `ChatInput.tsx`, `HeaderBar.tsx` |
| Utility/hook files | camelCase.ts | `sessionStorage.ts`, `useIsDarkMode.ts` |
| Transport files | kebab-case.ts | `websocket-client.ts` |
| Tool files | camelCase.ts | `getDocumentContent.ts` |
| Directories | lowercase | `components/`, `lib/`, `tools/` |
| Module constants | UPPER_SNAKE_CASE | `FALLBACK_MODELS`, `MAX_SESSIONS_PER_HOST` |
| Tool names (Office) | snake_case | `get_document_content`, `insert_table` |
| CSS-in-JS classes | camelCase | `inputContainer`, `messageUser` |
| Callback props | `on` prefix | `onSend`, `onChange`, `onModelChange` |

### TypeScript

- `interface` for object shapes (props, data models); `type` for unions and aliases
- Props interfaces are **not exported** (private to component file)
- Shared data interfaces **are exported**: `export interface Message { ... }`
- `strict: true` in tsconfig.json
- `catch (e: any)` is the standard pattern in this codebase (not `unknown`)
- Tool handler args are cast inline: `const { html } = args as { html: string }`

### React Patterns

- 100% functional components, arrow functions with `React.FC<Props>`:
  ```typescript
  export const ChatInput: React.FC<ChatInputProps> = ({ value, onChange }) => { ... };
  ```
- All public components use **named exports** (never `export default`)
- Styling via Fluent UI `makeStyles` — call `useStyles()` as first line in component
- State management: `useState` + prop drilling (no Redux, Context, Zustand)
- Custom hooks follow `use` prefix: `useIsDarkMode`, `useLocalStorage`
- `useCallback` for memoized callbacks; `useMemo` is not used

### File Organization (React components)

1. Imports
2. Exported interfaces/types (shared data models)
3. Non-exported interfaces (component props)
4. Module-level constants (UPPER_SNAKE_CASE)
5. `makeStyles` definition
6. Helper functions (non-exported)
7. Small sub-components (non-exported)
8. Main component (exported)

### Error Handling

Tool handlers return a structured error object on failure:
```typescript
{ textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} }
```

UI-layer errors are displayed as assistant messages. Silent catches (`catch {}`)
are acceptable for non-critical operations (session cleanup, localStorage reads).

Logging: `console.error` for errors, `console.log` with bracket-prefixed tags
(`[copilot-cli]`, `[tool.call]`), and `remoteLog()` for client-to-server error reporting.

### Async Patterns

- `async/await` everywhere — no `.then()` chains in TypeScript
- Fire-and-forget with `.catch(() => {})` for non-critical calls
- `for await (const event of session.query(...))` for streaming responses
- `AsyncGenerator` pattern in websocket-client.ts for event queuing

### Office.js Patterns

Every Office API interaction follows this structure:
```typescript
return await Word.run(async (context) => {
  const items = context.document.body.paragraphs;
  items.load(["text", "style"]);
  await context.sync();      // Must sync before reading loaded properties
  // ... work with items ...
  await context.sync();      // Final sync for writes
  return "Success message";
});
```

Key rules:
- Always use `X.run(async (context) => { ... })` (X = Word, Excel, PowerPoint)
- Always `.load()` properties explicitly before accessing them
- Always `await context.sync()` between load and read
- Multiple `context.sync()` calls per handler are normal
- Validate inputs before the `X.run()` call; validate indices after loading items

### Tool File Structure (Uniform)

Every tool in `src/ui/tools/` follows this exact pattern:
```typescript
import type { Tool } from "@github/copilot-sdk";

export const getDocumentContent: Tool = {
  name: "get_document_content",
  description: "...",
  parameters: {
    type: "object",
    properties: { ... },
    required: ["..."],
  },
  handler: async (args) => {
    try {
      return await Word.run(async (context) => { ... });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
```

Register new tools in `src/ui/tools/index.ts` by importing and adding to the
appropriate array (`wordTools`, `powerpointTools`, or `excelTools`).
