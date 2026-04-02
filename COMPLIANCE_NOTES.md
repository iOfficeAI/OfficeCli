# OfficeCLI Fork — Apache 2.0 Compliance Notes

## License Status
- Upstream: Apache License 2.0 (LICENSE file intact)
- Upstream NOTICE file: **none shipped** (checked v1.0.28)
- Apache 2.0 §4(b): modified files must carry prominent change notices

## Our Fork
- Fork owner: lidge-jun
- Fork URL: https://github.com/lidge-jun/OfficeCLI
- Upstream: https://github.com/iOfficeAI/OfficeCLI

## Planned Modifications
- CJK font handling (CjkHelper.cs) — Korean/Japanese/Chinese font metadata
- CJK language tag injection (w:lang, a:lang attributes)
- Kinsoku line-break processing
- East Asian character spacing

## Attribution (for distribution)
If we ship a bundled binary, include this NOTICE:

```
OfficeCLI
Copyright (c) iOfficeAI contributors

This product includes software developed at
iOfficeAI (https://github.com/iOfficeAI/OfficeCLI).

Licensed under the Apache License, Version 2.0.

Modifications by cli-jaw contributors:
- CJK (Korean/Japanese/Chinese) font handling and language tags
- CJK line-break (kinsoku) processing
- East Asian character spacing support
```

## Build Environment Note (Phase 02)
- .NET 10 SDK: **NOT INSTALLED** on this machine
- `dotnet build` and `dotnet publish`: DEFERRED to when .NET SDK is available
- Using pre-built release binary (v1.0.28) from GitHub Releases in the meantime
- Phase 03 (CJK integration) requires .NET SDK — install before proceeding
