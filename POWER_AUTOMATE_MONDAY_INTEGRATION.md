# Connecting Power Automate with Monday.com

This guide explains how to integrate Microsoft Power Automate with Monday.com to automate the sample calculator workflow.

## Table of Contents
1. [Overview](#overview)
2. [Prerequisites](#prerequisites)
3. [Setup Instructions](#setup-instructions)
4. [Integration Workflow](#integration-workflow)
5. [Example Scenarios](#example-scenarios)
6. [Troubleshooting](#troubleshooting)

## Overview

Power Automate (formerly Microsoft Flow) is a cloud-based automation service that allows you to create automated workflows between apps and services. Monday.com is a work operating system that enables teams to run projects and workflows with confidence. By connecting these two platforms, you can automate data flow between Monday.com and the Excel-based Sample Calculator.

### Integration Benefits
- **Automated Data Entry**: Automatically populate calculator inputs from Monday.com items
- **Real-time Updates**: Push calculation results back to Monday.com boards
- **Reduced Manual Work**: Eliminate copy-paste errors and save time
- **Audit Trail**: Maintain a record of all calculations in Monday.com

## Prerequisites

Before starting, ensure you have:

1. **Microsoft Account** with Power Automate access (included with Microsoft 365)
2. **Monday.com Account** with appropriate permissions
3. **Excel Online** with the Sample Calculator workbook
4. **Power Automate Connectors** for:
   - Monday.com (available in Power Automate)
   - Excel Online (Business)

## Setup Instructions

### Step 1: Prepare Your Excel Workbook

1. Upload your Sample Calculator workbook to OneDrive or SharePoint
2. Ensure the workbook contains the Office Script from `Sample7` (TypeScript version)
3. Note the file location - you'll need this path for Power Automate

### Step 2: Set Up Monday.com Board

1. Create a new board in Monday.com or use an existing one
2. Add the following columns to your board:
   - **Population** (Number column)
   - **Risk Level** (Dropdown: Low, Medium, High)
   - **Frequency** (Dropdown: Multiple per day, Daily, Weekly, Bi-Weekly, Monthly, Quarterly, Annually, As Needed)
   - **Sub Population** (Number column, optional)
   - **Base Sample Size** (Number column - for results)
   - **Sub Sample Size** (Number column - for results)

### Step 3: Connect Power Automate to Monday.com

1. Go to [Power Automate](https://make.powerautomate.com)
2. Click **Create** → **Automated cloud flow**
3. Name your flow (e.g., "Monday.com to Sample Calculator")
4. Search for the Monday.com trigger: **When an item is created**
5. Click **Create**

### Step 4: Configure Monday.com Trigger

1. Sign in to Monday.com when prompted
2. Select your **Board** from the dropdown
3. This trigger will fire whenever a new item is added to your Monday.com board

### Step 5: Add Excel Script Action

1. Click **+ New step**
2. Search for "Excel Online (Business)"
3. Select **Run script**
4. Configure the action:
   - **Location**: OneDrive or SharePoint
   - **Document Library**: Select where your workbook is stored
   - **File**: Browse and select your Sample Calculator workbook
   - **Script**: Select the script from Sample7 (it should appear as "main")

### Step 6: Map Input Parameters

Map the Monday.com column values to the Excel script parameters:

```
population: (Dynamic content) → Population
riskLevel: (Dynamic content) → Risk Level
frequency: (Dynamic content) → Frequency
subPop: (Dynamic content) → Sub Population (optional)
```

To add dynamic content:
1. Click in the parameter field
2. Select the appropriate Monday.com column from the dynamic content panel

### Step 7: Update Monday.com with Results

1. Click **+ New step**
2. Search for "Monday.com"
3. Select **Change column value**
4. Configure:
   - **Board**: Select your board
   - **Item ID**: (Dynamic content) → Item ID (from trigger)
   - **Column ID**: Select "Base Sample Size"
   - **Value**: (Dynamic content) → baseSampleSize (from Excel script)

5. Click **+ New step** again to add another update for Sub Sample Size:
   - **Board**: Select your board
   - **Item ID**: (Dynamic content) → Item ID (from trigger)
   - **Column ID**: Select "Sub Sample Size"
   - **Value**: (Dynamic content) → subSampleSize (from Excel script)

### Step 8: Save and Test

1. Click **Save** in the top right
2. Create a test item in your Monday.com board with sample data:
   - Population: 250
   - Risk Level: Medium
   - Frequency: Daily
3. Check that your flow runs successfully
4. Verify the calculated results appear in Monday.com

## Integration Workflow

The complete workflow operates as follows:

```
Monday.com Item Created
        ↓
Power Automate Triggered
        ↓
Extract Item Data (Population, Risk Level, Frequency, Sub Population)
        ↓
Call Excel Online Script (Sample7)
        ↓
Script Calculates Sample Sizes
        ↓
Return Results to Power Automate
        ↓
Update Monday.com Item with Results
```

## Example Scenarios

### Scenario 1: Audit Planning Automation

**Use Case**: Your audit team uses Monday.com to track upcoming audits. When a new audit is added, automatically calculate the required sample size.

**Setup**:
1. Create a Monday.com board called "Audit Schedule"
2. Add columns for audit parameters
3. Configure the Power Automate flow as described above
4. When auditors add new items, sample sizes are calculated automatically

### Scenario 2: Multi-Tier Auditing

**Use Case**: Calculate both base population and sub-population sample sizes for complex audits.

**Setup**:
1. Include both "Population" and "Sub Population" columns
2. Map both parameters in Power Automate
3. Both "Base Sample Size" and "Sub Sample Size" will be calculated and updated

### Scenario 3: Batch Processing

**Use Case**: Import multiple audit requirements at once from CSV or another system.

**Setup**:
1. Use Monday.com's import feature to bulk-add items
2. Power Automate will process each item automatically
3. All sample sizes will be calculated in sequence

### Scenario 4: Approval Workflow

**Use Case**: Require manager approval before finalizing sample sizes.

**Additional Steps**:
1. After calculating sample size, add a **Condition** step
2. Add an **Approval** action to request manager review
3. Update status column based on approval result

## Troubleshooting

### Issue: Flow not triggering

**Solutions**:
- Verify the Monday.com connection is active in Power Automate
- Check that the trigger is set to the correct board
- Ensure the Monday.com connector has necessary permissions
- Try reconnecting the Monday.com account

### Issue: Script execution fails

**Solutions**:
- Verify the Excel workbook is accessible in OneDrive/SharePoint
- Check that the Office Script is saved and named correctly
- Ensure all input parameters are being passed correctly
- Review script requirements in Sample7 file

### Issue: Invalid population or frequency values

**Solutions**:
- Ensure Monday.com dropdown values exactly match script expectations:
  - Risk Level: "Low", "Medium", or "High" (case-sensitive)
  - Frequency: Must match one of the valid options
- Add validation in Monday.com board settings
- Consider adding error handling in Power Automate with Condition steps

### Issue: Results not updating in Monday.com

**Solutions**:
- Check the "Change column value" action configuration
- Verify column names and IDs are correct
- Ensure the Item ID is being passed correctly
- Check Monday.com API rate limits

### Issue: Sub Sample Size not calculating

**Solutions**:
- Verify "Sub Population" value is being passed
- Check that the value is not empty or null
- The subPop parameter is optional - it only calculates if provided

## Advanced Configuration

### Adding Error Notifications

Add error handling to your flow:

1. After the "Run script" action, click **...** (three dots) → **Configure run after**
2. Select **has failed** and **has timed out**
3. Add a **Send an email** or **Post a message in Teams** action
4. Include error details from dynamic content

### Creating a Manual Trigger Option

To recalculate on-demand:

1. Create a new flow with trigger: **When an item is updated** (Monday.com)
2. Add a **Condition** to check if a specific column changed (e.g., "Recalculate" checkbox)
3. Follow the same script execution steps

### Logging and Monitoring

Track all calculations:

1. Add a **Compose** action after the script runs
2. Create a JSON object with input and output values
3. Add **Create item** (SharePoint) or **Add row** (Excel) to log the calculation

## Additional Resources

- [Power Automate Documentation](https://learn.microsoft.com/en-us/power-automate/)
- [Monday.com API Documentation](https://developer.monday.com/)
- [Office Scripts Documentation](https://learn.microsoft.com/en-us/office/dev/scripts/)
- [Sample Calculator Scripts](./Sample7) - TypeScript implementation

## Best Practices

1. **Test with sample data** before deploying to production boards
2. **Use descriptive names** for your flows and connections
3. **Document your workflow** in Monday.com board description
4. **Monitor flow runs** regularly in Power Automate dashboard
5. **Set up error notifications** to catch issues early
6. **Use Monday.com board templates** to standardize audit requests
7. **Implement version control** for Office Scripts via source control

## Support

For issues with:
- **Power Automate**: Contact Microsoft Support or visit the Power Automate community
- **Monday.com**: Contact Monday.com support or visit their developer community
- **Sample Calculator Scripts**: Review the documentation in this repository

---

**Last Updated**: February 2026
