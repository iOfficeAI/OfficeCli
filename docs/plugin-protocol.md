# OfficeCli Plugin Protocol

**Status**: Draft v0. Subject to change before v1 ratification.
**Audience**: Plugin authors and OfficeCli contributors.

## 1. Motivation

OfficeCli's main repo focuses on three universal Office formats (`.docx`, `.xlsx`,
`.pptx`). To extend format support without bloating the main binary or coupling
external implementations to the main repo's license, format support is delivered
through **plugins** — independent sidecar processes discovered and invoked by the
main binary.

Concrete drivers:

- Legacy formats (`.doc`, `.rtf`, `.odt`) where some users need migration but the
  parser is heavy and the format is fading
- Regional formats (`.hwpx`) maintained by communities outside the main team
- Export targets (`.pdf`, `.epub`) where the renderer library has size, license,
  or platform constraints that make in-tree bundling undesirable
- Proprietary implementations that need to stay out of the Apache-licensed main
  repo

## 2. Plugin Kinds

A plugin declares its **kind** in its manifest. Each kind has a fixed
responsibility, lifecycle, and IPC pattern. v1 defines three kinds.

### 2.1 `dump-reader` — read a foreign format, emit officecli commands

Used to **migrate** a foreign format into one of main's native formats
(`.docx`/`.xlsx`/`.pptx`). The output format is declared by the plugin's
manifest `target` field (default `"docx"`).

| Aspect | Value |
|---|---|
| Lifecycle | Short-lived (one shot) |
| Source file handle | Plugin (read-only) |
| Target file handle | Main (replays plugin's batch into a sibling native file) |
| Vocabulary | **Main's `<target>` command vocabulary** (no plugin-defined extensions) |
| IPC | None — plugin writes a `BatchItem[]` JSON array to stdout and exits |
| Output extension | Sibling `<source-stem>.<target>` next to the source |

Flow:

1. User invokes a command that opens a `.doc` file
2. Main checks for a sibling `<source-stem>.<target>` next to the source. If
   it exists and is newer than the source, main opens it directly and skips
   steps 3–5
3. Main spawns the plugin: `<plugin> dump <source>`
4. Plugin parses the source and prints a JSON array of `add`/`set`/`batch` items
   (same schema as `officecli batch --commands`) to stdout, then exits 0
5. Main creates a blank `<target>` skeleton, replays the batch against it, and
   moves it to the sibling path. Subsequent invocations reuse the sibling

Edits target the sibling native file, not the original source. Source-side changes
invalidate the cache automatically via mtime comparison; delete the sibling to
force reconversion.

### 2.2 `exporter` — convert native format to a foreign target

Used to **render** native content (`.docx`/`.xlsx`/`.pptx`) into a foreign output
file (e.g. `.pdf`). Single-direction, no editing.

| Aspect | Value |
|---|---|
| Lifecycle | Short-lived |
| Source file handle | Main (reads native file) |
| Target file handle | Plugin (writes foreign file) |
| Vocabulary | None — no commands exchanged |
| IPC | Optional; default is no pipe (just stdout/stderr + exit code) |

Flow:

1. User invokes a view mode that targets a foreign format (e.g.
   `officecli view <file> pdf --out <path>`). The mode name maps to the
   target extension.
2. Main resolves the `(from, to)` pair to a plugin
3. Main spawns the plugin with the source path and target path
4. Plugin reads the source (using its own libraries), writes the target
5. Plugin exits 0 if the target was written successfully

### 2.3 `format-handler` — own a foreign format end-to-end

Used to support a **first-class non-native format** (e.g. `.hwpx`). The plugin
holds the file open for the entire session and handles all document operations.

| Aspect | Value |
|---|---|
| Lifecycle | Long-lived (session duration) |
| Source file handle | Plugin (read-write, same file as target) |
| Target file handle | Same as source |
| Vocabulary | **Plugin-defined** (declared in manifest) |
| IPC | Required (pipe stays open) |

Flow:

1. User invokes a command on a `.hwpx` file
2. Main resolves `.hwpx` to a `format-handler` plugin
3. Main spawns the plugin with the file path and pipe name
4. Plugin opens the file and listens on the pipe
5. Main wraps the plugin in a `ProxyHandler : IDocumentHandler`; every operation
   (`add`/`set`/`get`/`query`/`save`/...) becomes an IPC message
6. On session end, main sends `close`; plugin exits

### 2.4 Reserved kinds

The following kinds are reserved for future use. Plugins MUST NOT declare them in
v1:

- `engine` — pluggable backend for an in-tree subsystem (e.g. PDF rendering,
  field refresh)

A plugin MAY declare multiple kinds in a single binary (e.g. an exporter that is
also a dump-reader). See §4.

## 3. Plugin Discovery

When main needs a plugin for `(kind, ext)`, it searches in this fixed order. The
first match wins.

1. **Environment variable**: `$OFFICECLI_PLUGIN_<KIND>_<EXT>` (absolute path to
   the plugin executable). Example: `$OFFICECLI_PLUGIN_DUMP_READER_DOC`.
2. **User plugins directory**:
   `~/.officecli/plugins/<kind>/<ext>/plugin(.exe)`
3. **Bundled plugins directory** (next to the main executable):
   `<dir>/plugins/<kind>/<ext>/plugin(.exe)`
4. **PATH lookup**: an executable named `officecli-<kind>-<ext>` or
   `officecli-<ext>` (in that priority).

Path conventions:

- `<kind>` uses kebab-case (`dump-reader`, `format-handler`, `exporter`)
- `<ext>` is the file extension without the leading dot (`doc`, `hwpx`, `pdf`)
- On Windows, `(.exe)` is appended automatically when searching
- Symlinks are followed

Main caches discovery results per process invocation. Adding a plugin between
invocations is picked up immediately.

## 4. Manifest

Every plugin MUST respond to `<plugin> --info` by printing a single JSON object
to stdout and exiting 0. The object describes the plugin to the main binary.

### 4.1 Required fields

| Field | Type | Description |
|---|---|---|
| `name` | string | Stable identifier, kebab-case (e.g. `officecli-doc`) |
| `version` | string | SemVer of the plugin (e.g. `1.0.0`) |
| `protocol` | integer | Protocol major version this plugin implements (currently `1`) |
| `kinds` | array | One or more declared kinds (see §2). Common case: `["dump-reader"]` |
| `extensions` | array | File extensions this plugin handles, leading dot (`[".doc"]`) |

### 4.2 Optional fields

| Field | Type | Description |
|---|---|---|
| `description` | string | Short human-readable description |
| `target` | string | Native format the plugin produces (`"docx"`/`"xlsx"`/`"pptx"`). Required-by-convention for `dump-reader`; defaults to `"docx"` if omitted. Determines the sibling-cache extension, the lint schema, and the handler that opens the replay output. |
| `tier` | string | Free-form tier identifier (`basic`/`pro`/`enterprise`) |
| `vocabulary` | object | Required for `format-handler`. See §4.3 |
| `supports` | array | Capability tags (e.g. `["tables","images","fields"]`) |
| `limits` | object | Plugin-imposed limits (e.g. `{"maxFileSizeMb": 200}`) |
| `homepage` | string | URL |
| `license` | string | SPDX identifier |

### 4.3 Vocabulary (format-handler only)

Format-handler plugins MUST declare the vocabulary their proxied document model
exposes:

```json
"vocabulary": {
  "addable_types": ["page", "annotation", "formfield", "outline-item"],
  "settable_props": {
    "annotation": ["type", "rect", "color", "contents", "author", "opacity"],
    "page": ["rotation", "mediaBox"],
    "formfield": ["value", "readOnly"]
  },
  "path_segments": ["/page[N]", "/page[N]/annotation[M]", "/formfield[<name>]"]
}
```

Main uses this for autocomplete, command validation, and help output. Main does
not interpret the semantics — it merely forwards commands using the declared
vocabulary.

### 4.4 Example manifests

`officecli-doc` (dump-reader):
```json
{
  "name": "officecli-doc",
  "version": "1.0.0",
  "protocol": 1,
  "kinds": ["dump-reader"],
  "extensions": [".doc"],
  "target": "docx",
  "tier": "basic",
  "supports": ["paragraphs", "runs", "tables", "images", "lists"]
}
```

`officecli-pdf-libreoffice` (exporter):
```json
{
  "name": "officecli-pdf-libreoffice",
  "version": "0.1.0",
  "protocol": 1,
  "kinds": ["exporter"],
  "extensions": [".pdf"],
  "supports": ["from:docx", "from:xlsx", "from:pptx"]
}
```

`officecli-hwpx` (format-handler):
```json
{
  "name": "officecli-hwpx",
  "version": "0.9.0",
  "protocol": 1,
  "kinds": ["format-handler"],
  "extensions": [".hwpx"],
  "vocabulary": {
    "addable_types": ["paragraph", "run", "table", "image", "footnote"],
    "settable_props": { ... },
    "path_segments": [ ... ]
  }
}
```

## 5. Invocation

Beyond `--info`, each kind has its own subcommand surface.

### 5.1 dump-reader

```
<plugin> dump <source-file> [--media-dir <dir>]
```

- `<source-file>`: absolute path to the file to read
- `--media-dir`: optional scratch directory the plugin may use for transient
  files (e.g. extracted images referenced by command paths)

Main sets the `OFFICECLI_BIN` environment variable to the path of the running
officecli binary, so plugins that produce an intermediate `.docx` (e.g. via an
external converter) can shell out to `officecli dump <converted.docx>` and pipe
its output to stdout instead of reimplementing docx → batch conversion. Plugins
that don't need this can ignore the variable.

The plugin writes a JSON array of batch commands to stdout. Schema matches
`officecli batch --commands`:

```json
[
  {"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": "Hello"}},
  {"command": "set", "path": "/body/paragraph[1]", "props": {"bold": "true"}}
]
```

Diagnostics go to stderr or `--log-file`. The plugin exits 0 on success; non-zero
codes follow §6.6.

### 5.2 exporter

```
<plugin> export <source-file> --out <target-file> [--options <json>]
```

- `<source-file>`: native format file (`.docx`/`.xlsx`/`.pptx`)
- `--out`: target path for the exported file
- `--options`: optional backend-specific options as a JSON string

### 5.3 format-handler

```
<plugin> open <file> --pipe <pipe-name>
```

The plugin opens the file, connects to the pipe (`\\.\pipe\<name>` on
Windows, the .NET cross-platform Unix-socket equivalent on Linux/macOS),
and serves messages until it receives `close` (see §6.4). Each request
gets exactly one reply before the next request is sent.

Request envelope (main → plugin):

```json
{"protocol":1, "msg_type":"command", "command":"<verb>", "args":{...}}
```

Proxied verbs:

**Read path:**

| `command` | `args` keys | `result` shape on `ok` |
|---|---|---|
| `view` | `mode` (`text`/`annotated`/`outline`/`stats`/`issues`), `start`/`end`/`max-lines`/`cols`/`type`/`limit`/`format` | string (or JSON object when `format=json`); for `mode=issues`, an array of issue objects |
| `get` | `path`, `depth` | DocumentNode JSON object |
| `query` | `selector` | array of DocumentNode |
| `validate` | (none) | array of `{errorType,description,path,part}` |

**Mutation path** (envelope carries `args` and `props` separately; `props` is
the user's `--prop key=value` dictionary, always string-to-string):

| `command` | `args` keys | `props` | `result` shape on `ok` |
|---|---|---|---|
| `set` | `path` | yes | array of string — names of props the plugin did **not** apply (`unsupported_property` UX) |
| `add` | `parent_path`, `type`, optional `position` | yes | string — path of the newly created element |
| `remove` | `path` | no | string or null — optional warning text (e.g. cells shifted) |
| `move` | `source_path`, optional `target_parent_path`, optional `position` | no | string — new path |
| `copy` | `source_path`, `target_parent_path`, optional `position` | no | string — new path |
| `raw` | `part_path`, optional `start_row`/`end_row`/`cols` | no | string — raw XML (or CSV-of-rows for spreadsheet parts) |
| `raw-set` | `part_path`, `xpath`, `action`, optional `xml` | no | null |
| `add-part` | `parent_part_path`, `part_type` | optional | object `{"rel_id":"...","part_path":"..."}` |
| `extract-binary` | `path`, `dest_path` | no | object `{"found":true,"content_type":"...","byte_count":N}` or `{"found":false}` |

`position` (when present) is `{"index":N}` OR `{"after":"<path>"}` OR
`{"before":"<path>"}` — at most one field set; all-null means append.

Reply envelopes (plugin → main):

```json
{"protocol":1, "msg_type":"ok", "result":<value>}
{"protocol":1, "msg_type":"error", "error":{"code":"<code>","message":"...","detail":"..."}}
```

A plugin that doesn't implement a given verb should reply with `error`
code `unsupported_command`; main translates this into the matching
user-facing error and (for `extract-binary` specifically) into a
`false` return so existing callers stay happy.

### 5.4 Universal options

All subcommands accept:

- `--log-file <path>`: append diagnostic output here instead of stderr
- `--quiet`: suppress non-error output

## 6. IPC Protocol

Only `format-handler` exchanges live messages with main; the framing below
applies to that kind. (`dump-reader` and `exporter` are short-lived and use the
simpler stdout / exit-code contracts described in §5.1 and §5.2.)

### 6.1 Transport

Named pipes (Windows) or Unix domain sockets (Linux/macOS). Main creates the
endpoint before spawning the plugin and passes the name via `--pipe`. The .NET
`NamedPipeServerStream` API provides this cross-platform.

### 6.2 Framing

UTF-8 text. One JSON object per line, terminated by `\n`. The protocol is
**request/response**: every client message receives exactly one server reply
before the next message is sent. For `format-handler`, **main is the client**
and **plugin is the server**.

### 6.3 Message envelope

Every message MUST include:

```json
{
  "protocol": 1,
  "msg_type": "<type>",
  ... type-specific fields ...
}
```

### 6.4 Message types

#### Request types (client → server)

| `msg_type` | Used by | Body |
|---|---|---|
| `command` | format-handler | `{ "command": "add"\|"set"\|..., "args": {...}, "props": {...} }` |
| `open` | format-handler | `{ "path": "<file>" }` (sent by main on session start) |
| `save` | format-handler | `{}` |
| `close` | format-handler | `{}` |
| `capabilities` | format-handler | `{}` (query what the server supports) |
| `ping` | format-handler | `{}` (liveness check) |

#### Response types (server → client)

| `msg_type` | Body |
|---|---|
| `ok` | `{ "result": <value-or-null> }` |
| `error` | `{ "error": { "code": "<code>", "message": "...", "detail": "..." } }` |
| `capabilities` | `{ "commands": [...], "vocabulary": {...} }` |

#### Server-pushed events (format-handler only)

| `msg_type` | Body |
|---|---|
| `event` | `{ "kind": "warning"\|"info", "message": "..." }` |

### 6.5 Error codes

Plugins SHOULD use these codes when applicable:

| Code | Meaning |
|---|---|
| `invalid_request` | Malformed message |
| `unsupported_command` | Recognized message but unimplemented |
| `unsupported_feature` | Recognized command but feature not in this build |
| `invalid_argument` | Argument failed validation |
| `not_found` | Target path/element does not exist |
| `corrupt_input` | Source file is malformed or unreadable |
| `license_expired` | Commercial plugin's license check failed |
| `internal_error` | Catch-all for plugin bugs |

Codes are extensible; main treats unknown codes as `internal_error`.

### 6.6 Exit codes

When a plugin process terminates:

| Code | Meaning |
|---|---|
| `0` | Success |
| `2` | Corrupt input file |
| `3` | Feature unsupported in this build |
| `4` | License expired |
| `5` | Protocol mismatch |
| `64`-`78` | Reserved (sysexits.h) |
| other | Plugin bug; main reports as `internal_error` |

## 7. Vocabulary Contract

### 7.1 Universal protocol shell (all kinds)

These elements are stable across all plugins and all kinds:

- Message envelope shape (§6.3)
- Command verbs: `add`, `set`, `remove`, `move`, `get`, `query`, `batch`,
  `raw-set`
- Path syntax: `/segment[N]` with `[N]` 1-based index OR `[<name>]` named
  reference
- Error code namespace (extensible)
- Exit code semantics

### 7.2 Per-format vocabulary

The specific **types** (`paragraph`/`page`/`cell`/...), **property names**
(`bold`/`fontsize`/`rect`/...), and **value formats** (`12pt`/`#FF0000`/...) are
not universal. They depend on which document model is at the other end:

- For `dump-reader`, the receiving model is main's `WordprocessingDocument`, so
  the vocabulary is main's docx vocabulary (published as
  `schemas/word-vocabulary.json`)
- For `format-handler`, the model is the plugin's own; the plugin declares its
  vocabulary in the manifest
- For `exporter`, there is no command vocabulary

## 8. Stability Commitments (Main → Plugins)

Once a protocol version is released, the main repository commits to:

1. **Protocol shell** is stable for the major version. Adding new optional
   message types is allowed; removing or changing types requires a major bump.
2. **docx vocabulary** (relevant to `dump-reader`): additions allowed; deletions
   or renames require a deprecation cycle of at least two minor releases with
   the old name accepted as an alias.
3. **Path syntax** does not change.
4. **Error code semantics** do not change. Adding new codes is allowed.
5. **Schema files** (`schemas/word-vocabulary.json`, etc.) are released
   alongside main and follow the same versioning.
6. **`capabilities` response** schema is forward-compatible: new fields may be
   added, existing fields stay.

## 9. Stability Commitments (Plugins → Main)

Plugin authors should:

1. Treat `--info` output schema as stable per protocol major version.
2. Implement graceful degradation when main lacks expected capabilities (query
   `capabilities` and skip unsupported features).
3. Provide a meaningful exit code on failure (don't silently exit 1 for every
   error).
4. Avoid writing to paths other than `--media-dir` and the declared output file.

## 10. Installation

The protocol does **not** mandate any installation mechanism. As long as the
plugin executable ends up at one of the discovery paths (§3), it works.

Common installation channels:

- **Manual**: download a release archive, extract to `~/.officecli/plugins/...`
- **Bundled distribution**: main's release archive includes a `plugins/`
  directory next to the executable
- **Built-in installer** (recommended for users): `officecli plugins install <name>`
- **Package managers**: `dotnet tool install`, `winget`, `brew`, `apt`, `scoop`
- **Enterprise deployment**: place binaries via IT distribution

The built-in installer consults a registry (default:
`https://officecli.ai/plugins/registry.json`; configurable for private mirrors)
which lists approved plugins, versions, download URLs, and SHA-256 hashes.

## 11. Writing a Plugin

### 11.1 Minimum dump-reader (C#)

```csharp
using System.Text.Json;

if (args[0] == "--info") {
    Console.WriteLine(JsonSerializer.Serialize(new {
        name = "officecli-doc-minimal",
        version = "0.0.1",
        protocol = 1,
        kinds = new[] { "dump-reader" },
        extensions = new[] { ".doc" }
    }));
    return 0;
}

// args: dump <source-file>
string sourcePath = args[1];

// Parse source file (your library here) and emit a batch on stdout:
var batch = new object[] {
    new {
        command = "add",
        parent = "/body",
        type = "paragraph",
        props = new { text = "Hello from .doc" }
    }
};
Console.WriteLine(JsonSerializer.Serialize(batch));
return 0;
```

### 11.2 Minimum exporter (Go)

```go
package main

import (
    "encoding/json"
    "fmt"
    "os"
    "os/exec"
)

func main() {
    if len(os.Args) > 1 && os.Args[1] == "--info" {
        json.NewEncoder(os.Stdout).Encode(map[string]interface{}{
            "name":       "officecli-pdf-min",
            "version":    "0.0.1",
            "protocol":   1,
            "kinds":      []string{"exporter"},
            "extensions": []string{".pdf"},
        })
        return
    }

    // args: export <source-file> --out <target-file>
    source := os.Args[2]
    var target string
    for i, a := range os.Args {
        if a == "--out" && i+1 < len(os.Args) {
            target = os.Args[i+1]
        }
    }

    cmd := exec.Command("soffice", "--headless", "--convert-to", "pdf",
        "--outdir", "/tmp/officecli-pdf", source)
    if err := cmd.Run(); err != nil {
        fmt.Fprintln(os.Stderr, err)
        os.Exit(3)
    }
    // ... move output to target ...
}
```

## 12. FAQ

**Q: Can plugins be in any language?**
A: Yes. The protocol is JSON over named pipes. Any language with subprocess and
pipe support works. .NET plugins can optionally use the `OfficeCli.Contracts`
NuGet package for type-safe types.

**Q: How does main know which plugin to use when several are installed?**
A: Discovery order (§3) is fixed and first-match-wins. For multiple installed
plugins for the same extension, users select via env var, config file, or
explicit `--plugin` flag.

**Q: Can a plugin be closed-source?**
A: Yes. Plugins are independent binaries with their own license. The protocol
is the only thing that's public.

**Q: Can a plugin be commercial?**
A: Yes. The plugin can do its own license check and exit 4 (`license_expired`)
when checks fail. The main repo's license does not propagate to plugin
implementations.

**Q: What if the plugin crashes?**
A: Main detects non-zero exit and surfaces a clear error to the user. Partial
state in main's in-memory document is discarded; no corrupt files are written.

**Q: What if the plugin hangs?**
A: Main applies a configurable timeout (default 60 s for short-lived kinds; no
timeout for format-handler sessions) and kills the plugin process.

**Q: What about pipes and security?**
A: Pipes are user-scoped on both Windows and Unix. Pipe names are randomized
per session. The plugin process runs with the same privileges as main.

**Q: How does this differ from MCP?**
A: MCP is for AI-tool interaction with officecli as a server. Plugins extend
officecli's format support; MCP exposes officecli to AI clients. The two are
complementary.

## 13. Versioning

This document tracks **protocol** version, distinct from main repo version.

- Protocol v1.x: Additive changes only (new optional fields, new message types,
  new error codes)
- Protocol v2.x: Breaking changes (removed/renamed fields, changed semantics)

Main repo declares supported protocol version(s) via `officecli --version`.
Plugins declare their target protocol in manifest. Main rejects plugins whose
major protocol version differs from any supported one.

## 14. Open Questions (pre-v1)

- Should `format-handler` plugins support concurrent multi-document sessions in
  one process? (v1: no, one process per open document)
- Should the registry support package signing? (Likely yes for v1.1)
- Should there be a `transformer` kind that takes one native format to another
  native format (e.g. `.docx → .pptx`)? (Deferred to v2)
- Should `capabilities` queries return JSON Schema fragments inline, or only
  list names? (Currently: names; consider inline schema in v1.1)

---

*This document is the source of truth for the OfficeCli Plugin Protocol. Changes
follow the deprecation rules in §8.*
