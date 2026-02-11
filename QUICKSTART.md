# Quick Start: Power Automate + Monday.com Integration

This is a condensed version of the full integration guide. For complete details, see [POWER_AUTOMATE_MONDAY_INTEGRATION.md](./POWER_AUTOMATE_MONDAY_INTEGRATION.md).

## 🎯 Goal
Automatically calculate audit sample sizes when new items are added to Monday.com.

## ⏱️ Time Required
15-20 minutes for initial setup

## 📋 Prerequisites Checklist
- [ ] Microsoft 365 account with Power Automate access
- [ ] Monday.com account
- [ ] Sample Calculator workbook uploaded to OneDrive/SharePoint
- [ ] Office Script (Sample7) added to the workbook

## 🚀 5-Minute Setup

### 1. Prepare Monday.com Board (3 minutes)

Create/configure a board with these columns:
- **Population** (Number)
- **Risk Level** (Dropdown: Low, Medium, High)
- **Frequency** (Dropdown: Daily, Weekly, Monthly, etc.)
- **Base Sample Size** (Number) - for results
- **Sub Sample Size** (Number) - for results (optional)

### 2. Create Power Automate Flow (5 minutes)

1. Go to https://make.powerautomate.com
2. Click **Create** → **Automated cloud flow**
3. Name it: "Monday.com Sample Calculator"
4. Add trigger: **Monday.com - When an item is created**
5. Select your board

### 3. Add Excel Script Action (3 minutes)

1. Add action: **Excel Online (Business) - Run script**
2. Select your workbook
3. Select the "main" script (from Sample7)
4. Map parameters:
   - `population` → Population (from Monday.com)
   - `riskLevel` → Risk Level (from Monday.com)
   - `frequency` → Frequency (from Monday.com)
   - `subPop` → Sub Population (from Monday.com, if applicable)

### 4. Update Monday.com with Results (4 minutes)

1. Add action: **Monday.com - Change column value**
   - Column: "Base Sample Size"
   - Value: `baseSampleSize` (from Excel script)

2. Add action: **Monday.com - Change column value**
   - Column: "Sub Sample Size"
   - Value: `subSampleSize` (from Excel script)

### 5. Test (5 minutes)

1. Save your flow
2. Create a test item in Monday.com:
   - Population: 250
   - Risk Level: Medium
   - Frequency: Daily
3. Check the flow run history
4. Verify results appear in Monday.com

## ✅ Success Criteria

You'll know it's working when:
- New Monday.com items automatically trigger the flow
- Base Sample Size is calculated and updated
- Flow run history shows successful executions
- No error notifications appear

## 🔧 Common Issues

| Issue | Quick Fix |
|-------|-----------|
| Flow not triggering | Reconnect Monday.com in Power Automate |
| Script not found | Ensure Sample7 script is saved in Excel Online |
| Wrong results | Check dropdown values match exactly (case-sensitive) |
| Missing sub-sample | Verify Sub Population has a value |

## 📚 Next Steps

After basic setup works:

1. **Add Error Handling**: Configure "run after" settings for failures
2. **Add Notifications**: Email or Teams alerts for errors
3. **Batch Processing**: Import multiple items via CSV
4. **Approval Workflow**: Add manager approval before finalizing

## 🆘 Need Help?

- **Full Documentation**: [POWER_AUTOMATE_MONDAY_INTEGRATION.md](./POWER_AUTOMATE_MONDAY_INTEGRATION.md)
- **Power Automate Help**: https://make.powerautomate.com
- **Monday.com Help**: https://support.monday.com

## 📊 Example Use Case

**Scenario**: Audit team planning quarterly reviews

1. Auditor creates Monday.com item:
   - Name: "Q1 Transaction Review"
   - Population: 500
   - Risk Level: High
   - Frequency: Multiple per day

2. Power Automate automatically:
   - Runs calculation
   - Returns: Base Sample Size = 60
   - Updates Monday.com item

3. Auditor sees result immediately and can proceed with sampling

---

**Time Saved**: ~5 minutes per calculation × multiple audits = hours saved per month!
