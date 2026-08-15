---
name: project_workflow_map
description: In-dashboard interactive Workflow map (how everything connects) with deep-links to each screen
metadata:
  type: project
  originSessionId: 23f44bb5-a672-44a4-9c2e-b8eac9975d80
---

Interactive **Workflow map** built into `dashboard.html` (commit 1183e68, 2026-06-25). A persistent **"🗺️ Workflow"** button sits in the shared `.main-header` (`#workflowBtn`) so it's reachable from every tab. Opens `#workflowOverlay` — a full-screen, color-laned visual of the whole lifecycle in 5 lanes: **A Lead & deal → B Build the event → C Run the money → D Amplify → ★ it all rolls up**. 18 steps, icons, minimal text, ⭐ flags the automated links (auto-file on sign, venue→capacity autofill, KPI roll-up).

Data lives in a JS array `WORKFLOW` (each step: lane, n, ico, t, where, short, detail, links[], optional `auto`). `renderWorkflowMap()` draws the bands; clicking a `.wf-step` → `showWfDetail(n)` shows a detail card; each detail has **"Open <screen> →"** buttons (`data-wf-go`) that call `activateTab(tab)` and close the map. Tab keys come from `module_registry` (inquiries/clients/agreements/events/venues/guests/income/expenses/sponsorships/social-calendar/strategy/actors). Steps that live INSIDE the event hub (auto-file, contracts, documents, tickets, donations) deep-link to **events** (open the event, then its hub tab). If a step's tab isn't enabled for the user's role, it toasts instead of no-op. Esc closes detail then map.

To extend: edit the `WORKFLOW` array + `WF_LANES`. The throwaway standalone `WORKFLOW_MAP.html` was removed — this in-app version supersedes it. The full end-to-end demo event used to validate the flow (slug `workflow-test-loft-party` + `(demo)` actors) was fully cleared afterward (DB + 18 storage objects). See [[project_crm_modules]] for the agreement→event auto-file that step 5 references.
