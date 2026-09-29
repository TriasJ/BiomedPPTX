# BiomedPPTX

## Quick Reference
- **Repo**: https://github.com/TriasJ/BiomedPPTX (branch: biomedpptx-main → main)
- **Solution**: `PowerPointLabs\PowerPointLabs.sln`
- **Version**: v0.2.1
- **Full plan/knowledge base**: `~/.claude/plans/dear-claude-i-want-buzzing-parnas.md`

## Critical Rules

### Never use AssemblyCatalog
MEF `AssemblyCatalog(Assembly.GetExecutingAssembly())` crashes because `System.Data.SQLite` types can't load. Always use `TypeCatalog` via `GetLoadableTypes()`. See plan for the pattern. Applied to: `BaseHandlerFactory`, `StylesDesigner`, `StyleOptionsFactory`, `StyleVariantsFactory`, `PictureSlidesLabWindowViewModel`.

### COM only for PowerPoint
All shape manipulation MUST use COM access (never python-pptx or Open XML SDK). VBA preferred for prototyping; C# VSTO for production.

### StyleCop C# 5 compatibility
Old StyleCop can't parse C# 6/7. No `$""`, `?.`, pattern matching, `out var`. Use `string.Format`, null checks, `as` casts. All braces required. Private after public. Usings alphabetical. One class per file.

### Push via gh
`git push` to github.com may timeout on this network. Use `gh` CLI or set `http.postBuffer 524288000`.

## Architecture

```
PowerPointLabs.sln
├── PowerPointLabs.csproj (VSTO add-in, .NET 4.7.2)
│   ├── SmartBrowserLab/     ── SMART-Library browser (4,474 illustrations, FTS5 search)
│   │   ├── Models/          ── SmartDatabase, BioArtFetcher, IllustrationItem, TilingMetadata
│   │   ├── Services/        ── ShapeInserter (PPTX copy → SVG → PNG → EMF fallback)
│   │   └── Views/           ── SmartBrowserPaneWPF (search, categories, tags, recolor)
│   ├── PatternBrushLab/     ── Pattern brush (238 tileable + custom patterns)
│   │   ├── Models/          ── PatternViewModel
│   │   ├── Services/        ── TilingEngine, PathExtractor, BrushModeController
│   │   └── Views/           ── PatternBrushPaneWPF (offset/angle/scatter/jitter/mode)
│   ├── ActionFramework/     ── MEF ribbon handlers
│   └── (all existing PPTLabs preserved)
├── Test.csproj
└── TestInterface.csproj
```

### Key Patterns
- **MEF handlers**: `[ExportActionRibbonId("Tag")]` — ribbon callback routing
- **Task panes**: WinForms UserControl + ElementHost → WPF UserControl
- **Clipboard**: `PPLClipboard.Instance.LockAndRelease()` for copy-paste (ShapesLab pattern)
- **Shape insertion**: PPLClipboard lock → copy → paste. Falls back to EMF export.
- **Source PPTX**: Opened as `Untitled: true` to prevent recovery on restart

## SMART-Library Assets
- `__Scratch/SMART-Library/` — illustrations.db (2.9MB SQLite+FTS5), png/ thumbnails, svg/ vectors
- `__Scratch/SMART-Lib/` — 49 source PPTX files (native PowerPoint grouped shapes)
- `~/.claude/skills/fetch-media/bioart_index.json` — BioArt catalog (661 items)

### Asset Path Resolution (ShapeInserter searches in order)
1. `<install dir>/Assets/SMART-Lib/` and `SMART-Library/`
2. `%APPDATA%/BiomedPPTX/Assets/`
3. `Documents/__Scratch/` (dev fallback)

## Critical Bugs & Solutions

| Bug | Solution |
|-----|----------|
| MEF ReflectionTypeLoadException | Always use TypeCatalog, never AssemblyCatalog |
| Clipboard copy fails (VSTO) | PPLClipboard.LockAndRelease → direct retry → EMF fallback |
| SMART PPTX recovery on restart | Open with Untitled:true, set Saved=true |
| PictureSlidesLab MahApps crash | Create WPF Application if null, load theme resources |
| Pattern selection lost on click | Only update _selectedPattern on non-null selections |
| PathExtractor offset bug | Use GenerateOutlinePath, not ConvertToFreeform |
| BioArt HTTPS fails | Set ServicePointManager.SecurityProtocol = Tls12 |

## Build

```bash
# Prerequisites: VS 2022 Community with Office/SharePoint workload
nuget restore PowerPointLabs/PowerPointLabs.sln -Source https://api.nuget.org/v3/index.json
msbuild PowerPointLabs/PowerPointLabs.sln /p:Configuration=Debug
```

### SQLite NuGet
- Package: `System.Data.SQLite.Core 1.0.118`
- DLL lives in `Stub.System.Data.SQLite.Core.NetFramework` (not the main package)
- Must import `.targets` for native interop DLL deployment (x86/ + x64/)

### Signing
- Self-signed `BiomedPPTX.pfx` (thumbprint: 435088C5...)

## Roadmap
1. **v0.3.0** — Installer + asset bundling (PRIORITY)
2. **v0.4.0** — Icons + UI polish
3. **v0.5.0** — Licensing + attribution
4. **v0.6.0** — Help + tutorial slides
5. **v0.7.0** — AI speech synthesis
