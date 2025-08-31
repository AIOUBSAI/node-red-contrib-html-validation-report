# node-red-contrib-html-validation-report

Two Node-RED nodes for validating workbook-like data and generating a rich HTML report.

- **data-validation-engine** (simplified): rule-based checks for:
  - `sheetsExist` — required sheet names present & non-empty
  - `sheetHasColumns` — required columns exist (supports dot paths)
  - Optional single condition per rule (attr/op/rhs)
  - Rules editable in the node or loaded from a JSON file under `userDir`
- **validation-report**: produces a modern, interactive HTML report:
  - Dark/light theme, sticky toolbar, search & highlight
  - Rule-level and sheet-level sections
  - Row pagination (10/25/50/100), counters, badges
  - Export visible rows (CSV, JSON), copy-to-clipboard, print

> **New in this repo**  
> Report node supports **typed input/output paths** (msg/flow/global).  
> It also handles logs that use `id` instead of `ruleId`, and `message` instead of `value`.

---

## Install

```bash
npm i node-red-contrib-html-validation-report
# then restart Node-RED
````

Or from Node-RED editor: **Menu → Manage palette → Install** → search for `node-red-contrib-html-validation-report`.

---

## Nodes

### 1) `data-validation-engine`

Validates an object model (usually built from Excel/CSV) using a small ruleset.

**Capabilities**

* Two rule types:

  * `sheetsExist`: `requiredSheets: string[]`
  * `sheetHasColumns`: `sheet: string`, `requiredColumns: string[]` (dot paths allowed)
* Optional rule-level condition:

  * `{ conditions: { and: [{ attribute, operator, rhsType, value }] } }`
  * Operators: `==`, `!=`, `contains`, `!contains`, `regex`, `isEmpty`, `!isEmpty`
  * RHS types: `str|num|bool|msg|flow|global|env|jsonata`
* Rules stored in-node **or** in a JSON file (`userDir`-relative), with lock & watch options

**Input**

* Source root chosen by **typed input**: `msg / flow / global`
* Path defaults to `msg.data`

**Output**

* `msg.validation = { logs, counts }`

  * `logs[]`: `{ id, type, level, message, description }`
  * `counts`: `{ info, warning, error, total }`
* Node status bubble shows `E: W: I:`

**Example rules JSON (file mode)**

```json
[
  {
    "type": "sheetsExist",
    "id": "RULE_SHEETS_EXIST",
    "description": "Verify that required sheets exist and are not empty",
    "requiredSheets": ["NAME", "PRICE"],
    "level": "error"
  },
  {
    "type": "sheetHasColumns",
    "id": "RULE_COLUMNS_NAME",
    "description": "Check that sheet includes name, etc fields",
    "sheet": "NAME",
    "requiredColumns": ["name", "grade" etc],
    "level": "warning"
  }
]
```

---

### 2) `validation-report`

Turns `msg.validation` (or any logs array) into a feature-rich HTML file.

**Input (typed)**

* Choose scope/path to read **validation** from (default: `msg.validation`)

  * Accepts either:

    * `{ logs, counts }`, or
    * `logs[]` directly

**Output (typed)**

* Choose scope/path to write **HTML** to (default: `msg.payload`)
* Optional **fixed filename** → written to chosen scope/path (default: `msg.filename`)

**Compatibility note**

* The report **groups** by `ruleId` **or** `id` (falls back to `rule` or `(unknown)`).
* The “Value” column uses `value` **or** `message` if `value` is absent.
  * The report maps `rules[].id → rules[].suggestions[]`.

**Typical flow**

```
[ inject ] → [ data-validation-engine ] → [ validation-report ] → [ file ]
```

* Configure **data-validation-engine** to read your model (e.g., `msg.data`).
* Configure **validation-report** input: `msg.validation`.
* Configure **validation-report** output: `msg.payload`.
* Set **Fixed filename** (optional): `logs/report.html`.
* File node writes `msg.payload` → `msg.filename`.

**Screenshot**
![Validation Report](docs/screenshot.png)

---

## Log format (expected)

Either:

```json
{
  "logs": [
    {
      "id": "RULE_SHEETS_EXIST",
      "type": "sheetsExist",
      "level": "error",
      "message": "Sheet 'PRICE' exists and is not empty.",
      "description": "Verify that required sheets exist and are not empty"
    }
  ],
  "counts": { "info": 0, "warning": 2, "error": 8, "total": 10 }
}
```

Or directly:

```json
[
  { "id": "RULE_SHEETS_EXIST", "type": "sheetsExist", "level": "error", "message": "..." },
  { "id": "RULE_COLUMNS_NAME", "type": "sheetHasColumns", "level": "warning", "message": "..." }
]
```

The report will display the rule header as the `id` (or `ruleId`), and use `message` as the “Value” when `value` is absent.

---

**File structure**

```
.
├─ data-validation-engine.js
├─ data-validation-engine.html
├─ validation-report.js
├─ validation-report.html
├─ package.json
├─ README.md
└─ docs/
   └─ screenshot.png
```

## License

MIT © AIOUBSAI
