# FM Tool Processor 🚀

Python replacement for the legacy Power Automate Desktop flow that:

1. Copies an `.xlsm` template to a processing folder.
2. Runs the `PopulateAndRunReport` macro with dynamic parameters.
3. Validates specific cells for SCAC and order/area sanity.
4. (Optionally) uploads the refreshed workbook to SharePoint.
5. Cleans up local files and returns a JSON result to the calling Flow.

The same codebase can:

* Run on a **Windows VM** (with Excel installed) for a drop-in replacement.
* Deploy as an **Azure Functions** HTTP endpoint (Windows Premium plan).

---

## 📁 Project Structure

See the tree in the commit or `/docs/architecture.md` (if you add one).  
Core logic lives in **`fm_tool_core/process_fm_tool.py`** and is imported
by the Azure Function wrapper in **`fm_tool_processor/`**.

---

## 🔧 Quick Start (Local CLI)

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
python -m fm_tool_core.process_fm_tool payload.json
```

## 📡 BID Webhook

Set `BID_WEBHOOK_URI` to the Power Automate webhook URL. When both
`BID-Payload` and `NOTIFY_EMAIL` are provided, the processor posts a
JSON payload to this endpoint to trigger downstream Flow steps.
