<!-- SPDX-License-Identifier: Apache-2.0 -->

# Lane B findings - COI gate & Office Dialog host

## 1. Office API used for two-way dialog messaging

The Office dialog channel is **string-only**. Two distinct APIs handle each direction:

| Direction | API |
|-----------|-----|
| dialog -> parent | `Office.context.ui.messageParent(str)` |
| parent -> dialog | `dialog.messageChild(str)` (task-pane side, Lane D) |
| dialog listens for parent | `Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, handler)` |
| parent listens for dialog | `Office.context.ui.addHandlerAsync(Office.EventType.DialogMessageReceived, handler)` (task-pane side, Lane D) |

### Exact names used in dialogHost.ts

```ts
// Register listener (dialog side):
Office.context.ui.addHandlerAsync(
  Office.EventType.DialogParentMessageReceived,
  handler,
  callback
);

// Send to parent (dialog side):
Office.context.ui.messageParent(encodeMessage(payload));
```

## 2. Requirement-set caveat - DialogApi 1.2

`DialogParentMessageReceived` and `messageChild()` (the symmetric parent->dialog
channel) require **DialogApi 1.2** (Office 2019 / M365 builds >= 1907; Excel on
the web since 2019).  DialogApi 1.1 has `displayDialogAsync` and the one-way
`DialogMessageReceived` event only.

Lane A (manifest owner) must add:

```xml
<Requirements>
  <Sets DefaultMinVersion="1.2">
    <Set Name="DialogApi" MinVersion="1.2"/>
  </Sets>
</Requirements>
```

Without this, `addHandlerAsync(DialogParentMessageReceived, ...)` silently no-ops
on older builds.

## 3. COEP mode - credentialless

DECISIONS.md #2 mandates `credentialless` (not `require-corp`). Pyodide (jsDelivr)
and the PyPI wheel do not send CORP headers; `require-corp` would block them.
`credentialless` is fully supported in Chromium-based hosts (WebView2/Edge).

## 4. COI service-worker belt-and-suspenders

`coi-serviceworker.js` is loaded as the first `<script>` in `dialog.html` so it
fires before Office.js and the module.  The SW's own `sessionStorage` guard
(`pcwCoiReloaded`) prevents reload loops - `coi.ts` intentionally adds no second
guard to avoid double-checking what the SW already handles.

## 5. Files delivered by Lane B

| File | Role |
|------|------|
| `src/dialog/coi.ts` | `ensureCrossOriginIsolated()` + `isolationDiagnostics()` |
| `src/dialog/dialog.html` | Dialog window page (COI host) |
| `src/dialog/dialogHost.ts` | Dialog entry point: COI check, runtime wiring, message dispatch |
| `docs/findings-lane-B.md` | This file |
