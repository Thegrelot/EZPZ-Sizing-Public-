# EZPZ Sizing — Nutanix Infrastructure Sizing Dashboard

A **local-first** desktop and web app that turns an **RVTools** or **Nutanix Collector** Excel export into an interactive sizing dashboard: KPIs, charts, OS mix, optional **PowerPoint** export, and a guided flow to create scenarios on **[sizer.nutanix.com](https://sizer.nutanix.com)** using your existing browser session.

---

## Table of contents

1. [Privacy and network use](#privacy-and-network-use)  
2. [What you can do in the app](#what-you-can-do-in-the-app)  
3. [Installation](#installation)  
4. [Data sources: RVTools vs Nutanix Collector](#data-sources-rvtools-vs-nutanix-collector)  
5. [Sidebar options](#sidebar-options)  
6. [Main dashboard: where each number comes from](#main-dashboard-where-each-number-comes-from)  
7. [Local sizing engine](#local-sizing-engine)  
8. [PowerPoint export](#powerpoint-export)  
9. [Nutanix Sizer integration (APIs and UI flow)](#nutanix-sizer-integration-apis-and-ui-flow)  
10. [Recent scenarios and “load scenario”](#recent-scenarios-and-load-scenario)  
11. [Building release binaries](#building-release-binaries)  
12. [Publishing a GitHub release](#publishing-a-github-release)  
13. [Project layout](#project-layout)  
14. [Terminology](#terminology)  
15. [Python dependencies](#python-dependencies)  

---

## Privacy and network use

| Activity | Network | What leaves this machine |
|----------|---------|---------------------------|
| Upload and parse `.xlsx` | **None** | Nothing — parsing is in-process with `pandas` / `openpyxl`. |
| Dashboard charts and KPIs | **None** | Plotly renders in the browser; data stays local. |
| PowerPoint generation | **None** | Charts are rasterized with **Kaleido** (local); slides are built with **python-pptx**. |
| **Create / open Nutanix Sizer scenarios** | **Yes** — `https://sizer.nutanix.com` | API calls authenticated with **your browser’s session cookie** (and a short-lived **Bearer JWT** for some search endpoints). Workload payloads you confirm in the wizard are sent like the official Sizer UI would. |

No telemetry or third-party analytics are built into this repository’s app code.

---

## What you can do in the app

- **Ingest** a VMware environment export as **RVTools `.xlsx`** or **Nutanix Collector `.xlsx`**.
- **Filter** by cluster or view **All** clusters.
- **Tune** analysis with toggles (powered-off VMs, stretched cluster, exclude common infra VMs).
- **Review** KPI cards, **Plotly** charts (storage, VM sizing, hosts, CPU models, OS breakdown, etc.).
- **Export** selected charts to a **Nutanix slate–style** `.pptx` deck.
- **Open the Sizer wizard** to search **Salesforce accounts / opportunities**, map each cluster to a **Sizer workload**, pick **hardware vendor**, **RF**, **utilization threshold**, **processor**, run **auto-size**, and get a **scenario URL** (`https://sizer.nutanix.com/scenario/S-{id}`).
- **List recent scenarios**, open one, and see a **sizing summary** table (nodes, model, utilization vs threshold, power, rack units, failover) when the Sizer API returns data.

---

## Installation

### Desktop (recommended)

Download the artifact from [GitHub Releases](https://github.com/Thegrelot/EZPZ-Sizing/releases).

| Platform | Artifact | Usage |
|----------|----------|--------|
| **macOS** | `RVToolsSizer-macOS.zip` | Unzip → open `RVToolsSizer.app` (first launch: right-click → **Open** if Gatekeeper blocks unsigned builds). |
| **Windows** | `RVToolsSizer-Windows.exe` | Run the executable; a splash window starts the embedded Streamlit server and opens the dashboard in your browser. |

### From source

```bash
git clone https://github.com/Thegrelot/EZPZ-Sizing.git
cd EZPZ-Sizing
pip install -r requirements.txt
streamlit run app.py
```

**Python 3.10+** is required.

---

## Data sources: RVTools vs Nutanix Collector

The sidebar radio switches the **parser pipeline**. Both paths normalize into the same internal column names so **`lib/sizing.py`** and the charts can share one code path.

### RVTools (`lib/parsing.py`)

**Export:** RVTools → **File → Export to Excel** → upload the `.xlsx`.

| Sheet | Role |
|-------|------|
| **vInfo** | VM inventory: name, power state, vCPU, memory, cluster, provisioned/in-use storage columns where present, OS strings. |
| **vDisk** | Per-VM disk capacity (supports cluster filter). |
| **vPartition** | Partition capacity / consumed (preferred for **consumed** storage rollups). |
| **vHost** | Physical hosts: cores, sockets, CPU model, CPU/memory usage % — used for **pCore** counts and host charts. |
| **vDatastore** | *(Optional)* Datastore capacity — used when present for datastore-oriented views; sizing still works if this sheet is missing. |

Column names are resolved with **case-insensitive** matching and common RVTools variants (e.g. `Capacity MiB` vs `Capacity MB`, alternate power-state / memory column labels).

### Nutanix Collector (`lib/collector_parsing.py`)

**Export:** Nutanix Collector tool → Excel — upload the `.xlsx`.

| Collector sheet | Maps to (conceptually) |
|-----------------|-------------------------|
| **vmList** | Same role as **vInfo** (VMs, power, vCPU, RAM, cluster, OS). |
| **vHosts** | Same role as **vHost**. |
| **vPartition** / **vDisk** | Same as RVTools partition/disk logic. |
| **vCluster**, **vCPU**, **vMemory** | Loaded when present for richer Collector exports; core sizing still relies on the main four sheets above. |

Collector **template** VMs are excluded from vmList-based analysis where applicable.

---

## Sidebar options

These apply per data source (separate session keys for RVTools vs Collector).

| Control | Effect |
|---------|--------|
| **Cluster** | `All` or a single cluster name — filters parsed `vInfo` / `vmList` and related joins before KPIs and charts. |
| **Include Powered-Off VMs** | When on, powered-off / suspended VMs are included in totals that normally use only **powered-on** VMs. |
| **Stretched Cluster** | Models a **site failure** by effectively using about **half** the hosts per cluster for pCore / ratio math (ratio rises accordingly). |
| **Exclude Infra VMs** | Drops known infrastructure patterns from workload totals (e.g. vCLS, Nutanix CVM, vCenter/VCSA, NSX, HCX, SRM, vROPS, etc.) — see in-app help for the full intent. |

---

## Main dashboard: where each number comes from

Unless a toggle says otherwise, **powered-on VMs** drive vCPU/RAM and most VM-centric charts. **Storage** prefers **consumed** (from **vPartition**) then **provisioned** fallbacks (partition → vDisk / vInfo), consistent with `compute_sizing()` in `lib/sizing.py`.

### KPI row (typical mapping)

| KPI | Primary data |
|-----|----------------|
| Total **vCPU** | Sum of `CPUs` on in-scope powered-on rows (`vInfo` / `vmList`). |
| Total **RAM** | Sum of `Memory_MiB` → converted to TB for display. |
| **vCPU : pCore** | Total vCPU ÷ total physical cores from **vHost** / **vHosts** in scope. Color bands: healthy / warning / aggressive ratio (server-style thresholds in the UI). |
| **Storage consumed / provisioned** | Aggregated from partition/disk/VM columns as above. |
| **Largest VM (vCPU / RAM)** | Max powered-on VM by CPUs and by memory. |
| **NUMA hint** | Compares largest VM vCPU to **cores-per-socket** on the largest host — flags when the VM may span sockets. |

### Charts (high level)

Examples include: storage efficiency (consumed vs provisioned), VM count by vCPU and RAM buckets, resources by cluster, power state donut, top VMs by vCPU/RAM, storage by cluster, host counts, pCores per cluster, hosts by CPU model, cores-per-host distribution, CPU model × cluster table with avg usage, **OS breakdown** (Windows Server vs Linux/Other vs Unknown using both OS columns where available).

Exact chart titles and small layout tweaks may evolve between releases; the **data contracts** are the parsed frames above.

---

## Local sizing engine

Implemented in **`lib/sizing.py`** (`compute_sizing` and helpers such as `resources_by_cluster`, `cluster_workloads_for_sizer`, bucket/top-N helpers).

- **Workload type** internally supports server vs BCA-style ratio highlighting where the UI wires it.  
- **N+1**: after removing the largest host, remaining cores and an adjusted ratio can be surfaced for resilience context.  
- **OS labels**: regex-based classification on RVTools-style and Collector-style OS strings (`_os_label_series`).  
- **Sizer handoff**: `cluster_workloads_for_sizer()` builds the per-cluster aggregates the wizard sends as **RAW_INPUT / Cluster Sizing** style payloads.

---

## PowerPoint export

- Charts expose **include-in-deck** checkboxes; the sidebar **Generate / Export PowerPoint** builds a `.pptx` from **`templates/slate_template.pptx`** (placeholders replaced with PNG snapshots of selected figures/tables).
- **Kaleido** (`plotly`) renders figures to PNG. On **macOS**, the app may batch exports via subprocesses (`launcher.py`) for speed and stability.
- **Optional environment variables** (for operators tuning export time vs resolution): e.g. `RVToolsSizer_PPTX_HQ=1` for larger default PNG dimensions; `RVToolsSizer_KALEIDO_WORKERS` / `RVToolsSizer_KALEIDO_CHUNK_TARGET` for parallel Kaleido behavior on Darwin — see `app.py` / `launcher.py` for current defaults.

---

## Nutanix Sizer integration (APIs and UI flow)

All calls go to **`https://sizer.nutanix.com`** via **`lib/sizer_api.py`**, with cookies from **`lib/browser_auth.py`**.

### Authentication

1. **Session cookie** — Read from a local browser profile (Firefox, Chrome, Edge, Brave, Chromium, Opera) for the `sizer.nutanix.com` domain. You must already be logged into Sizer in that browser.  
2. **Bearer JWT** — Obtained with **`GET /api/sizer/v1.0/segpt/token`** using that session; cached until shortly before expiry. Used for endpoints that reject cookie-only access (notably **scenario search**).

### Wizard → API mapping

| User step / feature | HTTP (relative to `https://sizer.nutanix.com`) | Notes |
|---------------------|-----------------------------------------------|--------|
| Account search (2+ chars) | **`GET /v1/accounts/search/?searchkey=`** | Session cookie. |
| Opportunities for account | **`GET /v1/scenario-search/accounts/{accountId}/opportunity`** | Session cookie; returns Sizer `opp_Id` used when creating a scenario. |
| Create scenario | **`POST /v3/scenarios`** | Session cookie; body includes name, SFDC account, opportunity id, portfolio version, client build string, etc. |
| Add workloads (one per cluster row) | **`POST /v1/design/{envId}/workloads`** | Session cookie; one request per workload JSON the UI built. |
| Auto-size (Nutanix pass) | **`POST /v1/environment/{envId}/auto-size`** | Session cookie; may send `{}` or a full vendor/threshold body depending on options (see `auto_size()`). |
| Auto-size (OEM / second vendor pass) | Same **`/v1/environment/{envId}/auto-size`** | Separate POST(s) with vendor-specific body from `_vendor_autosize_body()`. |
| Wait until sizing finishes | Poll **`GET /v3/scenarios/{envId}/sizing-status`** | Then read **`GET /v3/scenarios/{envId}/sizing-summary`** to discover **`clusterId`** values for follow-up auto-size. |
| Processor dropdown / CPU matching | **`get_processor_list()`** | Uses an **embedded static catalogue** extracted from Sizer UI (not a live HTTP catalogue endpoint). **`match_cpu_to_sizer_processor()`** maps **vHost** CPU model strings to catalogue entries for defaults. |

Returned **`envId`** is the internal scenario id; the shareable link uses **`S-{scenario_id}`** (`scenario_url()`).

---

## Recent scenarios and “load scenario”

| Feature | APIs / behavior |
|---------|------------------|
| **Recent scenarios** table | **`POST /api/sizer/v1.0/scenarios/search`** with Bearer JWT — body requests “created by me”, ordered by last update. Search box filters **client-side** (server errors if server-side `filterRules` are sent). Fallback candidates in code include legacy **`GET /v1/scenario-search/scenarios`** etc. |
| **Sizing summary** (per cluster: nodes, model, util %, thresholds, kW, RU, N+x) | **`GET /v3/scenarios/{envId}/sizing-summary`** + **`GET /v3/scenarios/{envId}/hardware-summary`** + workload fetch for N+ metadata — merged in **`get_sizing_summary()`**. |
| **Reload workload table from Sizer** | **`get_scenario_workloads()`** tries multiple paths (e.g. **`/v3/scenarios/{id}/workloads`**, **`/v1/design/{id}/workloads`**, **`/api/sizer/v1.0/scenarios/{id}/workloads`**) with cookie and/or Bearer until a workload list parses. |
| **Threshold label** when loading a scenario | **`get_scenario_threshold()`** reads **`/v3/scenarios/{id}/sizing-summary`** utilization thresholds and maps to **Default (85%)** / **Consolidate (95%)**. |
| **Vendor per cluster** when loading | **`get_scenario_vendor_map()`** reads **`/v3/scenarios/{id}/hardware-summary`** (`currentRecommendations` / `leadTimeData`). |

Developer-only probing of alternate endpoints may exist (`probe_scenario_endpoints_debug`) for troubleshooting API shape changes.

---

## Building release binaries

```bash
pip install -r requirements-build.txt
python build.py
```

- **macOS:** produces `dist-macos/RVToolsSizer.app` (relocatable venv bundle by default). Optional **`python build.py --mac-frozen`** for a PyInstaller-frozen variant (may affect Kaleido behavior).  
- **Windows:** produces `dist-windows/RVToolsSizer.exe`.  

Zip the `.app` for upload:

```bash
cd dist-macos && zip -r RVToolsSizer-macOS.zip RVToolsSizer.app
```

`build.py` looks for **`static/logo.png`** (or `static/nutanix-icon.png`) to build a **`.icns`** on macOS when Pillow and `iconutil` succeed.

---

## Publishing a GitHub release

```bash
git tag v1.0.2 -m "Release v1.0.2"
git push origin v1.0.2

gh release create v1.0.2 \
  --title "EZPZ Sizing v1.0.2" \
  --notes "Short release notes here." \
  "dist-macos/RVToolsSizer-macOS.zip#RVToolsSizer-macOS.zip" \
  "dist-windows/RVToolsSizer.exe#RVToolsSizer-Windows.exe"

# Add another artifact later
gh release upload v1.0.2 "dist-windows/RVToolsSizer.exe#RVToolsSizer-Windows.exe"
```

Use **`refs/for/...`** if your remote is Gerrit; for GitHub, push tags to **`origin`** as above.

---

## Project layout

```
EZPZ-Sizing/
├── app.py                  # Streamlit UI, charts, PPTX orchestration, Sizer wizard
├── launcher.py             # Desktop splash, Streamlit worker, Kaleido CLI/batch helper
├── build.py / build.sh / build.bat
├── requirements.txt / requirements-build.txt
├── static/                 # App + window icon (e.g. logo.png)
├── templates/              # slate_template.pptx
└── lib/
    ├── parsing.py          # RVTools XLSX
    ├── collector_parsing.py
    ├── sizing.py           # Local sizing math and OS classification
    ├── sizer_api.py        # sizer.nutanix.com client
    ├── browser_auth.py     # Cookies + JWT for Sizer
    ├── pptx_export.py      # Placeholder / fallback images
    └── …
```

---

## Terminology

| Term | Meaning |
|------|---------|
| **pCore** | Physical CPU core on an ESXi host. |
| **vCPU** | Virtual CPUs assigned to a VM. |
| **vCPU : pCore** | Oversubscription — how many vCPUs share each physical core. |
| **RF2 / RF3** | AOS container **replication factor** (copies of data). |
| **N+1** | Capacity after losing **one** host (or N+x as reported by Sizer). |
| **NUMA** | VM vCPU count vs socket core count — large VMs may need socket-aware placement. |
| **CVM** | Controller VM — Nutanix per-node service VM (often excluded from “customer” workload toggles). |
| **envId** | Sizer internal scenario / environment id used in REST paths. |

---

## Python dependencies

| Package | Role |
|---------|------|
| `streamlit` | Web UI |
| `pandas` / `openpyxl` | XLSX ingest |
| `plotly` / `kaleido` | Charts and PNG export for PPTX |
| `python-pptx` | PowerPoint assembly |
| `Pillow` | Icons / image helpers in launcher and build |
| `requests` | Sizer HTTPS API |
| `browser-cookie3` | Read browser cookie stores for `sizer.nutanix.com` |

See **`requirements.txt`** for pinned minimum versions.

---

## License and support

This project is maintained as an **independent tooling** layer on top of public-facing Nutanix Sizer behavior. Sizer UI and API responses can change; if a release breaks scenario load or search, compare browser DevTools traffic with **`lib/sizer_api.py`** and update the client accordingly.

For questions or issues, use **GitHub Issues** on this repository.
