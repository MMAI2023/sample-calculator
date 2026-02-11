# Connecting Power Automate to Monday.com

This guide walks through the first step of connecting Monday.com to Power Automate and building a simple analysis flow using the sample calculator.

---

## Prerequisites

- A [Microsoft Power Automate](https://flow.microsoft.com) account (included with Microsoft 365)
- A [Monday.com](https://monday.com) account with at least one board
- A Monday.com API token (see [Step 1](#step-1-obtain-your-mondaycom-api-token))

---

## Step 1: Obtain Your Monday.com API Token

1. Log in to your Monday.com account.
2. Click your **profile avatar** (bottom-left corner).
3. Select **Administration** → **Connections** → **API**.
4. Under **Personal API Token**, click **Generate** (or **Copy** if one already exists).
5. Save the token securely — you will use it in Power Automate.

> **Tip:** Treat this token like a password. Do not commit it to source control.

---

## Step 2: Create a New Power Automate Flow

1. Open [Power Automate](https://flow.microsoft.com).
2. Click **+ Create** → **Instant cloud flow**.
3. Give your flow a name, for example: `Monday.com Sample Analysis`.
4. Choose the trigger **Manually trigger a flow** and click **Create**.

---

## Step 3: Add an HTTP Action to Connect to Monday.com

Monday.com uses a GraphQL API. Power Automate connects to it through the built-in **HTTP** action.

1. Click **+ New step** → search for **HTTP** → select the **HTTP** action.
2. Configure the action as follows:

| Field | Value |
|-------|-------|
| **Method** | `POST` |
| **URI** | `https://api.monday.com/v2` |
| **Headers** | `Content-Type`: `application/json` |
| | `Authorization`: `<your API token>` |
| **Body** | See the query below |

### Example Body — Fetch Board Items

```json
{
  "query": "{ boards(ids: [YOUR_BOARD_ID]) { name items_page { items { id name column_values { id text } } } } }"
}
```

Replace `YOUR_BOARD_ID` with the numeric ID of your Monday.com board. You can find it in the board URL: `https://your-org.monday.com/boards/1234567890` → the ID is `1234567890`.

3. Click **Save** and then **Test** to confirm the connection returns data.

---

## Step 4: Parse the Monday.com Response

1. Click **+ New step** → search for **Parse JSON** → select the action.
2. Set **Content** to the `Body` output of the HTTP step.
3. Click **Generate from sample** and paste a sample response from Step 3's test run, or use the schema below:

```json
{
  "type": "object",
  "properties": {
    "data": {
      "type": "object",
      "properties": {
        "boards": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "name": { "type": "string" },
              "items_page": {
                "type": "object",
                "properties": {
                  "items": {
                    "type": "array",
                    "items": {
                      "type": "object",
                      "properties": {
                        "id": { "type": "string" },
                        "name": { "type": "string" },
                        "column_values": {
                          "type": "array",
                          "items": {
                            "type": "object",
                            "properties": {
                              "id": { "type": "string" },
                              "text": { "type": "string" }
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
}
```

---

## Simple Analysis: Sample Size Calculation from Monday.com Data

Once connected, you can build a flow that reads population and risk data from a Monday.com board and runs it through the sample calculator.

### Monday.com Board Setup

Create a board with the following columns:

| Column | Type | Purpose |
|--------|------|---------|
| **Name** | Default | Audit area name |
| **Population** | Numbers | Total population count |
| **Risk Level** | Status | Low / Medium / High |
| **Frequency** | Status | Daily, Weekly, Monthly, etc. |
| **Sample Size** | Numbers | Result (written back by the flow) |

### Flow Overview

```
┌─────────────────────────────────────┐
│  Trigger: Manually trigger a flow   │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  HTTP: Fetch items from Monday.com  │
│  (GraphQL POST to api.monday.com)   │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Parse JSON: Extract board items    │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Apply to each: Loop over items     │
│  ┌────────────────────────────────┐ │
│  │ Run script: Execute Sample7    │ │
│  │ (Excel Online / Office Script) │ │
│  │ Inputs: population, risk,      │ │
│  │         frequency              │ │
│  └─────────────┬──────────────────┘ │
│                │                    │
│                ▼                    │
│  ┌────────────────────────────────┐ │
│  │ HTTP: Write sample size back   │ │
│  │ to Monday.com via mutation     │ │
│  └────────────────────────────────┘ │
└─────────────────────────────────────┘
```

### Write Results Back to Monday.com

To write the calculated sample size back to Monday.com, add another **HTTP** action inside the loop:

| Field | Value |
|-------|-------|
| **Method** | `POST` |
| **URI** | `https://api.monday.com/v2` |
| **Headers** | Same as Step 3 |
| **Body** | See the mutation below |

```json
{
  "query": "mutation { change_simple_column_value(item_id: ITEM_ID, board_id: BOARD_ID, column_id: \"sample_size\", value: \"CALCULATED_VALUE\") { id } }"
}
```

Replace `ITEM_ID`, `BOARD_ID`, and `CALCULATED_VALUE` with dynamic values from the loop.

---

## Summary

| Step | Action | Purpose |
|------|--------|---------|
| 1 | Obtain API token | Authenticate with Monday.com |
| 2 | Create flow | Set up the Power Automate flow |
| 3 | HTTP action | Connect to Monday.com GraphQL API |
| 4 | Parse JSON | Extract board data |
| 5 | Run Sample7 script | Calculate sample sizes |
| 6 | HTTP mutation | Write results back to Monday.com |

This gives you an end-to-end integration: data flows from Monday.com into the sample calculator and results are written back automatically.
