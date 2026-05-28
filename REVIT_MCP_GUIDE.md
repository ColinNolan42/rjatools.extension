# Revit MCP Server — Installation, Usage & Session Constraints

## Overview

The Revit MCP (Model Context Protocol) server creates a live bridge between Claude and an open Revit model. Claude can read model data, query elements, and execute C# code directly inside Revit — all from a chat session. No manual copy-paste or scripting required.

**Architecture:**
```
Claude (chat) ←→ MCP Server (Node.js / npm) ←→ RevitMCPPlugin.dll (Revit addin) ←→ Revit Model
```

---

## Installation Guide

### Prerequisites

| Requirement | Version |
|---|---|
| Autodesk Revit | 2025 (plugin compiled for 2025) |
| Node.js | 18+ |
| npm | 9+ |
| Claude Desktop or Claude Code (VS Code extension) | Latest |

---

### Step 1 — Install the MCP npm Package

Open a terminal (PowerShell or Command Prompt) and run:

```powershell
npm install -g mcp-server-for-revit
```

Verify installation:
```powershell
mcp-server-for-revit --version
```

The package installs to:
```
C:\Users\[USERNAME]\AppData\Roaming\npm\node_modules\mcp-server-for-revit\
```

---

### Step 2 — Install the Revit Addin

1. Download or locate the addin files. The required folder structure is:

```
C:\Users\[USERNAME]\AppData\Roaming\Autodesk\Revit\Addins\2025\
├── mcp-servers-for-revit.addin          ← addin manifest
└── revit_mcp_plugin\
    ├── RevitMCPPlugin.dll                ← main plugin
    └── Commands\
        ├── commandRegistry.json          ← command definitions
        └── RevitMCPCommandSet\
            └── 2025\
                └── RevitMCPCommandSet.dll
```

2. The `.addin` manifest content should be:
```xml
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>mcp-servers-for-revit</Name>
    <Assembly>revit_mcp_plugin/RevitMCPPlugin.dll</Assembly>
    <AddInId>090a4c8c-61dc-426d-87df-e4bae0f80ec1</AddInId>
    <FullClassName>revit_mcp_plugin.Core.Application</FullClassName>
    <VendorId>mcp-servers-for-revit</VendorId>
    <VendorDescription>https://github.com/mcp-servers-for-revit</VendorDescription>
  </AddIn>
</RevitAddIns>
```

3. Restart Revit. The plugin loads silently on startup — no ribbon button appears.

---

### Step 3 — Configure Claude Desktop

Open (or create) the Claude Desktop config file:
```
C:\Users\[USERNAME]\AppData\Roaming\Claude\claude_desktop_config.json
```

Add the MCP server entry:
```json
{
  "mcpServers": {
    "mcp-server-for-revit": {
      "command": "node",
      "args": [
        "C:/Users/[USERNAME]/AppData/Roaming/npm/node_modules/mcp-server-for-revit/build/index.js"
      ]
    }
  }
}
```

> Replace `[USERNAME]` with your Windows username. Use forward slashes in the path.

Restart Claude Desktop after saving.

---

### Step 4 — Configure Claude Code (VS Code Extension)

If using the Claude Code VS Code extension instead of Claude Desktop, add to your project's `.claude/settings.json` or user settings:

```json
{
  "mcpServers": {
    "mcp-server-for-revit": {
      "command": "node",
      "args": [
        "C:/Users/[USERNAME]/AppData/Roaming/npm/node_modules/mcp-server-for-revit/build/index.js"
      ]
    }
  }
}
```

---

### Step 5 — Verify Connection

1. Open Revit and load a model
2. Start a Claude chat session
3. Ask Claude: **"run say_hello"**
4. A dialog should appear in Revit saying "Hello MCP!"
5. If it works, the connection is live

**If connection fails:**
- Confirm Revit is open with a model loaded
- Check the addin loaded: Revit → Add-Ins tab → External Tools (should list mcp-servers-for-revit)
- Restart the MCP server node process if needed
- Recheck the path in `claude_desktop_config.json`

---

## Available Tools

Once connected, Claude has access to the following MCP tools:

| Tool | Description |
|---|---|
| `say_hello` | Test connection — shows dialog in Revit |
| `send_code_to_revit` | Execute arbitrary C# code inside Revit |
| `get_current_view_info` | Info about the active view |
| `get_current_view_elements` | All elements in the active view |
| `get_selected_elements` | Currently selected elements |
| `get_available_family_types` | All family types in the model |
| `ai_element_filter` | Query elements by criteria |
| `operate_element` | Select, hide, isolate elements |
| `color_elements` | Color elements by parameter |
| `create_point_based_element` | Place point-based families |
| `create_line_based_element` | Place line-based elements |
| `delete_element` | Delete by ElementId |
| `export_room_data` | Extract rooms with properties |
| `get_material_quantities` | Material takeoffs |
| `analyze_model_statistics` | Element counts by category/type |
| `create_grid` | Generate grid system |
| `create_room` | Place rooms |
| `create_level` | Create levels |
| `create_dimensions` | Add dimension annotations |
| `store_project_data` | Store key/value data in model |
| `query_stored_data` | Retrieve stored data |

---

## Using `send_code_to_revit`

This is the most powerful tool. It executes C# code directly inside Revit.

### Code Template

Your code runs inside this template:
```csharp
public static class CodeExecutor {
    public static object Execute(Document document, object[] parameters) {
        // YOUR CODE RUNS HERE
        // Return a string to get output back in Claude
    }
}
```

### Critical Rules

| Rule | Detail |
|---|---|
| Variable name | `document` (lowercase d) — NOT `doc`, `Document`, or `uiApp` |
| Transactions | **Do NOT create a `Transaction`** — the server wraps execution automatically. Using `new Transaction(...)` or `using (var t = ...)` will throw an exception. Just call `param.Set(value)` directly. |
| Return value | Always `return "some string"` — Claude reads this as output |
| Namespaces | `Autodesk.Revit.DB.*` is available. `System.Collections.Generic`, `System.Linq` are available. |
| Linked models | Use `RevitLinkInstance.GetLinkDocument()` to access linked files |

### Read-Only Example
```csharp
var sheets = new FilteredElementCollector(document)
    .OfClass(typeof(ViewSheet))
    .Cast<ViewSheet>()
    .OrderBy(s => s.SheetNumber)
    .Select(s => s.SheetNumber + " - " + s.Name)
    .ToList();
return string.Join("\n", sheets);
```

### Write Example (No Transaction Needed)
```csharp
var sheet = new FilteredElementCollector(document)
    .OfClass(typeof(ViewSheet))
    .Cast<ViewSheet>()
    .FirstOrDefault(s => s.SheetNumber == "A-101");

var vp = document.GetElement(sheet.GetAllViewports().First()) as Viewport;
var p = vp.LookupParameter("Title on Sheet");
p.Set("NEW TITLE");
return "Done: " + p.AsString();
```

### Querying Linked Models
```csharp
var link = new FilteredElementCollector(document)
    .OfClass(typeof(RevitLinkInstance))
    .Cast<RevitLinkInstance>()
    .FirstOrDefault(l => l.Name.Contains("Arch"));

var linkedDoc = link.GetLinkDocument();

var rooms = new FilteredElementCollector(linkedDoc)
    .OfCategory(BuiltInCategory.OST_Rooms)
    .WhereElementIsNotElementType()
    .ToList();

return "Room count: " + rooms.Count;
```

---

## Known Issues & Tips

| Issue | Solution |
|---|---|
| `using (var t = new Transaction(...))` crashes | Never use `using` blocks with Transaction — server already handles it |
| `Document` (capital D) fails | Use `document` (lowercase) — it's the template variable name |
| Connection drops | Revit was closed or restarted — run `say_hello` to confirm live before writes |
| Workshared model — "reserved workset" error | Another user owns that element; changes won't save. Try element by element to isolate conflicts. |
| Empty result from view-filtered collector | Fixtures/elements may be in a linked model, not the host |
| `RevitLinkInstance.GetLinkDocument()` returns null | Linked file is unloaded — reload it in Manage Links |
| Chinese error text `执行失败` | General execution failure — simplify code, check for null references |

---

## Testing Period Constraints for Claude

> **These rules apply during the testing period and must be followed in every session.**

### Before Any Write Operation

Claude **must**:

1. **Present a plan** explaining:
   - What elements will be changed
   - What the current values are
   - What the new values will be
   - How many elements are affected

2. **Ask for explicit approval** before executing any `send_code_to_revit` call that modifies the model (sets parameters, creates elements, deletes elements, renames views, etc.)

3. **Report results** after execution — confirm what changed, how many succeeded, any skipped/failed.

### Read Operations

Read-only queries (`FilteredElementCollector`, `GetParameter`, `AsString`, etc.) **do not require approval** — Claude may run these freely to gather information.

### Example Approved Workflow

```
Claude:   "I will update the Title on Sheet for 3 viewports:
           - ADM M-101 [1]: 'ADMIN PLAN' → 'MECHANICAL PLAN - FIRST FLOOR'
           - ADM M-102 [1]: 'ADMIN PLAN' → 'MECHANICAL PLAN - SECOND FLOOR'
           - ADM M-103 [1]: 'ADMIN PLAN' → 'MECHANICAL PLAN - ROOF'
           Shall I proceed?"

User:     "Yes" / "Change M-102 to X instead" / "Skip M-103"

Claude:   [executes] "Done. 3 updated, 0 skipped."
```

### Operations Requiring Approval

- Setting any parameter value
- Creating elements (views, sheets, dimensions, annotations, rooms)
- Deleting elements
- Renaming views or sheets
- Modifying family instances
- Any batch operation affecting more than 1 element

---

## Session Startup Checklist

When starting a new session involving Revit MCP:

1. Ask Claude to run `say_hello` to confirm connection
2. If working on a specific project, tell Claude:
   - Which model is open
   - What building/discipline you're working on
   - Any constraints (worksharing, linked models, etc.)
3. Remind Claude of the testing period constraints if needed

---

*Last updated: 2026-05-27 — RJA Tools / Fraser PW Project*
