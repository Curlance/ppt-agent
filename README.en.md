<p align="center"><img src="assets/logo.png" width="112" alt="ppt-agent logo" /></p>
<h1 align="center">ppt-agent</h1>
<p align="center"><strong>Let agents work in real PowerPoint. See every step.</strong></p>
<p align="center">Live reading and editing · Visible windows · Activity feed · MCP / HTTP / CLI</p>
<p align="center"><a href="README.md">简体中文</a> · <strong>English</strong></p>
<p align="center">
  <a href="https://github.com/Curlance/ppt-agent/actions/workflows/ci.yml"><img src="https://github.com/Curlance/ppt-agent/actions/workflows/ci.yml/badge.svg" alt="CI" /></a>
  <img src="https://img.shields.io/badge/version-0.1.0-orange" alt="Version 0.1.0" />
  <img src="https://img.shields.io/badge/Python-3.11%2B-3776AB?logo=python&logoColor=white" alt="Python 3.11+" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D4" alt="Windows" />
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-22A699" alt="MIT" /></a>
</p>
<p align="center"><a href="#features">Features</a> · <a href="#quick-start">Quick start</a> · <a href="#screenshots">Screenshots</a> · <a href="#agent-integration">Integration</a> · <a href="#validation-and-limits">Validation</a></p>

<picture><source media="(max-width: 680px)" srcset="assets/workflow-mobile.svg" /><img src="assets/workflow.svg" alt="Agent → serial COM → visible PowerPoint, with live preview and activity feed" width="100%" /></picture>

ppt-agent controls the Windows desktop version of PowerPoint in real time. Agents inspect open presentations, edit text and shapes, update notes and table cells, and verify changes through readback and screenshots. PowerPoint follows the target slide and shape while a browser observer displays the live slide and a Chinese activity feed.

> **Status: a usable core under development.** Live PowerPoint reading, editing, saving, previews and MCP transports have been tested locally. Office add-ins and the DSH status bar still need installation and end-to-end validation inside their hosts. See the [validation report (Chinese)](交付报告.md).

## Features

| Capability | Implemented behavior |
|---|---|
| Read | Presentations, slides, text, notes, shapes, table contents and grouped shapes |
| Edit | Text and formatting, positions, textboxes, images, notes and table cells |
| Slides | Add, duplicate, delete and navigate; follow the target slide and shape |
| Files | Open, create, save, back up and close a specific presentation |
| Visibility | Visible PowerPoint windows, live preview, activity feed, pacing and stop controls |
| Interfaces | 24 tools defined once and exposed through MCP, HTTP/OpenAPI and CLI |
| COM ownership | One daemon and one STA worker, validated inputs, no automatic write replay |

## Quick start

Requires **Windows, Python 3.11+, and an installed, launchable desktop Microsoft PowerPoint**. Local validation used Python 3.13 and PowerPoint 16.0. PowerPoint for the web does not expose the COM interface used here.

Run in PowerShell; replace 3.13 with your installed supported Python version if needed:

```powershell
git clone https://github.com/Curlance/ppt-agent.git
cd ppt-agent
py -3.13 -m venv .venv
.venv\Scripts\python.exe -m pip install -e ".[dev]"
.venv\Scripts\python.exe -m pptd doctor
.venv\Scripts\python.exe -m pptd url --launch
```

The doctor checks the interpreter and COM, starting the daemon when needed. The second command opens the observer with its local token. Open and inspect a deck:

```powershell
.venv\Scripts\python.exe -m pptd open "C:\path\to\demo.pptx"
.venv\Scripts\python.exe -m pptd call ppt_read_slide -a slide=1
.venv\Scripts\python.exe -m pptd call ppt_shot -a slide=1
.venv\Scripts\python.exe -m pptd call ppt_set_pace -a pace_ms=800
```

Use the shape index or `ref` returned by inspection, then verify the edit:

```powershell
# Replace shape=1 with the actual identifier returned by inspection.
.venv\Scripts\python.exe -m pptd call ppt_set_text -a slide=1 -a shape=1 -a text="New title"
.venv\Scripts\python.exe -m pptd call ppt_read_slide -a slide=1
```

`python -m pptd tools` lists the tools. `python -m pptd stop` stops the daemon and leaves PowerPoint open. Use the environment's Python executable for these commands.

## Screenshots

These are actual browser screenshots with a presentation created by the UI tests.

**Observer** — live slide, PowerPoint state, tools and activity feed.

![Observer](docs/screenshots/observer.png)

**Narrow task pane page** — browser rendering validated; loading inside PowerPoint is still pending.

<p align="center"><img src="docs/screenshots/taskpane.png" width="380" alt="Actual narrow task pane page screenshot" /></p>

## Agent integration

Installation provides `pptctl` inside the environment. It is equivalent to `.venv\Scripts\python.exe -m pptd`.

| Transport | Entry point | Client |
|---|---|---|
| MCP stdio | `python -m pptd mcp` | Local MCP clients; configure interpreter, working directory and arguments |
| MCP HTTP | `http://127.0.0.1:8792/mcp` | Streamable HTTP clients with a token |
| HTTP / OpenAPI | `http://127.0.0.1:8791` | `/tools/<name>` or `/call` |
| CLI | `pptctl call <name>` | Scripts, terminals and local agents |

All transports reuse the daemon. The token lives in `%LOCALAPPDATA%\ppt-agent\runtime.json`; `pptctl url --mcp-token` displays it. Keep it out of shared configuration and screenshots.

See the [connection guide](docs/api/CONNECT.md), [tool reference](docs/api/tools.md) and [OpenAPI](docs/api/openapi.json). The public guide uses placeholder paths. Export local configuration with:

```powershell
.venv\Scripts\python.exe -m pptd export-api --local-paths --out local-docs
```

- **DSH:** `pptctl dsh-bundle` generates a bundle with local paths; `pptctl dsh-status` checks installation and loading.
- **Agent skill:** [skills/ppt-agent](skills/ppt-agent/SKILL.md) describes inspection, editing, verification and recovery. `pptctl skill-install` defaults to DSH; use `--out` for another directory.
- **Office add-ins:** [installation instructions](addin/README.md) cover the VBA window and Web Add-in; installation and trust configuration are separate steps.

## Validation and limits

```powershell
# No real PowerPoint required
.venv\Scripts\python.exe -m pytest -q
.venv\Scripts\python.exe -m pptd export-api --check

# Launches PowerPoint and creates temporary presentations
.venv\Scripts\python.exe -m pytest -m com -q
```

Windows CI checks unit tests, generated documentation and packaging. Desktop Office is absent from CI; real COM and browser tests run locally. See the [validation report](交付报告.md) for the recorded scope.

| Limit | Current scope |
|---|---|
| Add-ins | Sources, configuration and pages are provided; in-host interaction still needs validation |
| Charts, animation, masters | Chart type/title reading only; complete data editing, arbitrary animation and master design are not implemented |
| Writes after timeout | An executing COM call cannot be forcibly rolled back; inspect the current state before proceeding |
| Recovery | PowerPoint determines undo granularity; use `ppt_save(mode="copy")` before major edits |
| Compatibility | Local validation does not establish compatibility with every Office version and environment |

## Layout

```text
pptd/               Daemon, STA COM, registry, tools and transports
addin/              VBA sources, Web Add-in manifest and instructions
skills/ppt-agent/   Agent workflow and verification rules
docs/api/           Generated contracts and connection guide
docs/screenshots/   Real UI screenshots of the generated demo deck
assets/             Logo and workflow diagrams
tests/              Unit, real COM and browser tests
```

[Design notes (Chinese)](项目介绍.md) preserve the early design; current completion status is documented here and in the validation report. Runtime records, tokens, private keys, local path configurations, backups and build artifacts are excluded from Git.

Licensed under [MIT](LICENSE). See [brand notes](assets/BRAND.md) for the logo prompt. This project is not affiliated with Microsoft.
