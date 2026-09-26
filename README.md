# BiomedPPTX

**Biomedical PowerPoint Extensions** — A fork of [PowerPointLabs](https://github.com/PowerPointLabs/PowerPointLabs) with integrated biomedical illustration tools.

## Features

### SMART-Library Browser
- **4,474 searchable biomedical illustrations** from the SMART Medical Illustrations library
- 49 categories (Cell membrane, Receptors, Blood, Nervous system, etc.), 10 tag filters
- Full-text search (FTS5/LIKE) by name, description, and topic
- Insert as **editable grouped shape** (primary), SVG, or PNG
- **NCBI BioArt** online toggle — 661 additional illustrations from NIH, SVG-first with lazy disk caching
- **Hue/Saturation recolor** — adjust colors of inserted grouped shapes while preserving whites/blacks/grays
- Right-click context menu: **Use as Path Pattern (1D)** / **Fill Pattern (2D)** sends directly to Pattern Brush

### Pattern Brush
- **238 tileable biological patterns** (phospholipid bilayers, epithelial cells, muscle fibers, vascular stents, etc.)
- **Custom patterns** from any SMART-Library illustration
- **1D path tiling**: tiles along lines, freeforms, and **shape outlines** (circles, triangles, stars, hexagons, diamonds, crosses, and all polygons)
- **2D grid tiling**: brick pattern with automatic offset rows for tissue layers
- **Mode selector**: Auto / Along Outline / Fill Area
- Controls: **Offset X/Y** (-50 to +100pt), **Angle** (-180 to +180deg), **Scatter**, **Jitter**
- **Draw** button for quick guide line creation
- **Brush Mode** for auto-convert on drawing

### PowerPointLabs (All Features Preserved)
All original PowerPointLabs features are intact:
- **AnimationLab** — Auto-animate, drill-down zoom, step-back
- **ColorsLab** — Color picker and palette tools
- **ShapesLab** — Shape gallery and management
- **EffectsLab** — Spotlight, blur, magnify, transparency
- **PictureSlidesLab** — Picture-based slide design
- **CropLab** — Crop to shape, slide, aspect ratio
- **PasteLab** — Paste at original position, fill slide
- **PositionsLab** — Align, distribute, swap shapes
- **ResizeLab** — Resize and match shapes
- **SyncLab** — Synchronize formatting
- **TimerLab** — Presentation timer
- **AgendaLab** — Auto-generate agenda slides
- **NarrationsLab** — Text-to-speech narrations
- **CaptionsLab** — Auto-generate captions
- **HighlightLab** — Highlight bullets and text
- **ZoomLab** — Zoom to area
- **e-Learning Lab** — Create interactive e-learning content

## Build Requirements

- **Visual Studio 2022 Community** with "Office/SharePoint development" workload
- **.NET Framework 4.7.2**
- NuGet restore required (includes `System.Data.SQLite.Core`)

```bash
# Clone
git clone https://github.com/TriasJ/BiomedPPTX.git
cd BiomedPPTX

# Open in Visual Studio
# PowerPointLabs/PowerPointLabs.sln

# NuGet restore + Build
# Or from command line:
nuget restore PowerPointLabs/PowerPointLabs.sln -Source https://api.nuget.org/v3/index.json
msbuild PowerPointLabs/PowerPointLabs.sln /p:Configuration=Debug
```

## Asset Setup

Place SMART-Library assets in one of these locations (searched in order):

1. `<add-in install dir>\Assets\SMART-Library\` + `SMART-Lib\`
2. `%APPDATA%\BiomedPPTX\Assets\SMART-Library\` + `SMART-Lib\`
3. `Documents\__Scratch\SMART-Library\` + `SMART-Lib\`

Required files:
- `SMART-Library/illustrations.db` — FTS5-searchable SQLite database (2.9 MB)
- `SMART-Library/png/` — Thumbnail images for browser grid
- `SMART-Lib/*.pptx` — 49 source PPTX files for editable shape insertion

## Acknowledgements

- **PowerPointLabs** — Original add-in by School of Computing, National University of Singapore ([GPLv2](LICENSE))
- **SMART Medical Illustrations** — Servier Medical Art, available under [CC BY 3.0](https://creativecommons.org/licenses/by/3.0/)
- **NCBI BioArt** — NIAID Visual & Medical Arts, Public Domain / CC-BY

## License

BiomedPPTX is released under GPLv2 (same as PowerPointLabs).
