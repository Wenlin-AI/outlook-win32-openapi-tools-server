# outlook-win32-openapi-tools-server

> **Spec‑first, Windows‑native OpenAPI service for automating Microsoft Outlook Tasks through the Win32 COM object model.**

---

## Overview

`outlook-win32-openapi-tools-server` turns the classic desktop Outlook client into a micro‑service: it exposes a small set of Task‑management endpoints (list, create, complete, delete) described by an OpenAPI 3.1 contract and backed by the **Win32 COM** object model accessed with [`pywin32`](https://pypi.org/project/pywin32/).\
This lets AI agents, scripts, and no‑code platforms automate personal or shared *Tasks* without OAuth, Graph tokens, or Internet connectivity—perfect for on‑prem, air‑gapped, or heavily locked‑down Windows environments.

|            | Win32 COM (this repo)                  | Microsoft Graph (sibling repo) |
| ---------- | -------------------------------------- | ------------------------------ |
| Dependency | Outlook desktop app                    | Microsoft 365 cloud tenant     |
| Auth mode  | Current Windows/Outlook session        | OAuth 2.0 / Azure AD           |
| Offline?   | ✅ Works offline                        | ❌ Requires network             |
| Repo       | **outlook-win32-openapi-tools-server** | ms-graph-openapi-tools-server  |

---

## Why another tool server?

1. **Agent‑ready interface** – By shipping an openly documented spec at `/openapi.json`, any LLM or orchestration framework that understands OpenAPI can reason about available operations and invoke them directly.
2. **No cloud dependency** – Many corporates prevent Graph API calls. Win32 COM works wherever Outlook works.
3. **Scriptable tasks** – Replace brittle VBA macros with a modern, container‑friendly micro‑service you can start via `uvicorn` or Docker.

---

## Features

- **OpenAPI 3.1** spec served at `/openapi.json`; interactive Swagger UI at `/docs`.
- **List / Create / Complete / Delete** task endpoints (more coming—reminders, categories, recurrence…).
- **Folder‑aware** – Point the server at any Tasks folder (Inbox Tasks, Teams channel, PST archive) via CLI flag or `OUTLOOK_TASKS_FOLDER` env var.
- **CLI wrapper** – Launch with `outlook-win32-openapi-tools-server --folder "\\Personal Folders\\My Tasks"`.
- **Packaging** – Managed by [`uv`](https://github.com/astral-sh/uv) in `pyproject.toml`; publishable to PyPI.
- **Docker** – Windows‑based container (`mcr.microsoft.com/windows/servercore:ltsc2022`) for CI/auto‑scaling scenarios.
- **Example integrations** – LangChain, Open WebUI, Zapier CLI.

---

## Quick Start (Local)

### 1. Prerequisites

- Windows 10/11 with \*\*Outlook Classic \*\*installed and configured.
- Python 3.11+.

### 2. Install

```bash
pip install git+https://github.com/your-org/outlook-win32-openapi-tools-server.git
```

### 3. Run

```bash
# Default folder = default Tasks folder in default mailbox
outlook-win32-openapi-tools-server --host 0.0.0.0 --port 8124 \
  --folder "\\Personal Folders\\My Tasks"
```

Then open [http://localhost:8124/docs](http://localhost:8124/docs) to try the API.

---

## Configuration Reference

| Method | Key                    | Example                                         | Description                                         |
| ------ | ---------------------- | ----------------------------------------------- | --------------------------------------------------- |
| CLI    | `--folder`             | `"\\Personal Folders\\Project X\\Action Items"` | MAPI path of the Tasks folder to operate on.        |
| Env    | `OUTLOOK_TASKS_FOLDER` | same as above                                   | Same purpose when CLI not provided.                 |
| Env    | `OUTLOOK_PROFILE`      | `"Contoso"`                                     | Force a specific Outlook profile if multiple exist. |
| Env    | `HOST`                 | `0.0.0.0`                                       | Host address to bind the server to.                 |
| Env    | `PORT`                 | `8124`                                          | Port to run the server on.                          |
| Env    | `LOG_LEVEL`            | `INFO`                                          | Standard FastAPI/uvicorn logging level.             |

### Environment File

The application supports loading environment variables from a `.env` file in the project root. Copy the `.env-example` file to `.env` and customize as needed:

```bash
cp .env-example .env
# Then edit .env with your preferred settings
```

---

## API Surface (v0.1)

| Method & Path                    | Summary                                               |
| -------------------------------- | ----------------------------------------------------- |
| `GET /tasks`                     | List all incomplete tasks in the configured folder.   |
| `POST /tasks`                    | Add a new task. Body: `{ subject, dueDate?, body? }`. |
| `POST /tasks/{taskId}/complete` | Mark task complete.                                   |
| `DELETE /tasks/{taskId}`        | Permanently delete task.                              |

(See **`openapi.json`** for full schema. Each listed task includes a `taskId` for referencing.)

### Example: create a task

```bash
curl -X POST http://localhost:8124/tasks \
     -H "Content-Type: application/json" \
     -d '{"subject":"Submit Q2 budget","dueDate":"2025-07-01"}'
```

---

## Docker (Windows containers)

```bash
# Pull the pre‑built image published by GitHub Actions (GitHub Container Registry)
docker pull ghcr.io/your-org/outlook-win32-openapi-tools-server:latest

# Run (Windows host)
docker run --rm -p 8124:8124 \
  -e OUTLOOK_TASKS_FOLDER="\Personal Folders\My Tasks" \
  ghcr.io/your-org/outlook-win32-openapi-tools-server:latest
```

> **Note:** If you prefer to build locally, clone the repo and run `docker build -t outlook-win32-openapi-tools-server .`. Linux containers cannot access Win32 COM—use Windows containers or run the server directly on the host.

---

## Directory Layout

```
.
├── app/
│   ├── main.py          # FastAPI application factory
│   ├── outlook.py       # Thin Win32 COM wrapper (connect, helpers)
│   └── routers/
│       └── tasks.py     # CRUD routes
├── docs/                # Project documentation & ADRs
├── examples/            # Integration snippets (LangChain, WebUI…)
├── openapi.json         # Auto‑generated API schema from FastAPI
├── Dockerfile           # Windows Server Core image
├── pyproject.toml       # Dependencies + uv entry point
└── README.md            # This file
```

---

## Examples

### LangChain Agent

```python
from langchain.tools import OpenAPITool
from langchain.agents import initialize_agent

tool = OpenAPITool.from_openapi_url(
    name="outlook-tasks",
    url="http://localhost:8124/openapi.json",
)
agent = initialize_agent([tool], llm=your_llm, agent="zero-shot-react-description")
print(agent.run("Add a task: Order toner cartridges by next Friday"))
```

### Open WebUI

```yaml
# examples/openwebui.yaml
tools:
  - name: outlook-tasks
    spec_url: http://localhost:8124/openapi.json
    base_url: http://localhost:8124
```

---

## Contributing

Bug reports & PRs welcome! Please read `docs/requirements.md` and follow the commit & branching guidelines inherited from the [example-openapi-tools-server](https://github.com/Wenlin-AI/example-openapi-tools-server/) template.

---

## License

MIT © Henri Wenlin and contributors

