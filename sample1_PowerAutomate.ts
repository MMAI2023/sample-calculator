/**
 * BASE SAMPLE CALCULATOR - Power Automate Office Script
 *
 * This Office Script can be called directly from Power Automate
 * to calculate the base sample size for audit sampling.
 *
 * It writes inputs to the worksheet cells, calculates the result,
 * and writes the output back — keeping everything aligned.
 *
 * WORKSHEET CELL MAPPING:
 *   C4  = Population (input)
 *   C6  = Risk Level (input)
 *   C8  = Frequency (input)
 *   C10 = Base Sample Size (output)
 *   B12 = Status/Error message (output)
 *
 * INPUTS (function parameters):
 *   population : number  - The actual population size (minimum 1)
 *   riskLevel  : string  - "Low", "Medium", or "High"
 *   frequency  : string  - "As Needed", "Multiple per day", "Daily", "Weekly",
 *                           "Bi-Weekly", "Monthly", "Quarterly", "Annually"
 *
 * OUTPUTS (returned object):
 *   sampleSize : number  - The calculated base sample size (0 if error)
 *   error      : string  - Error message (empty string if successful)
 */
function main(
  workbook: ExcelScript.Workbook,
  population: number,
  riskLevel: string,
  frequency: string
): { sampleSize: number; error: string } {
  // Get the worksheet
  let ws = workbook.getWorksheet("Base Sample Calculator");

  // Write inputs to cells for alignment
  if (ws) {
    ws.getRange("C4").setValue(population);
    ws.getRange("C6").setValue(riskLevel);
    ws.getRange("C8").setValue(frequency);
  }

  // Validate inputs
  if (population < 1) {
    const errorMsg = "Please enter a valid population (minimum 1)";
    if (ws) { ws.getRange("B12").setValue(errorMsg); ws.getRange("C10").setValue(""); }
    return { sampleSize: 0, error: errorMsg };
  }
  if (!["Low", "Medium", "High"].includes(riskLevel)) {
    const errorMsg = "Please select a valid risk level (Low, Medium, or High)";
    if (ws) { ws.getRange("B12").setValue(errorMsg); ws.getRange("C10").setValue(""); }
    return { sampleSize: 0, error: errorMsg };
  }
  const validFrequencies = [
    "As Needed", "Multiple per day", "Daily", "Weekly",
    "Bi-Weekly", "Monthly", "Quarterly", "Annually"
  ];
  if (!validFrequencies.includes(frequency)) {
    const errorMsg = "Please select a valid frequency";
    if (ws) { ws.getRange("B12").setValue(errorMsg); ws.getRange("C10").setValue(""); }
    return { sampleSize: 0, error: errorMsg };
  }

  // Calculate sample size
  const result = calculateSampleSize(population, riskLevel, frequency);

  // Write outputs to cells for alignment
  if (ws) {
    if (result.error) {
      ws.getRange("B12").setValue(result.error);
      ws.getRange("C10").setValue("");
    } else {
      ws.getRange("C10").setValue(result.sampleSize);
      ws.getRange("B12").setValue("");
    }
  }

  return result;
}

function calculateSampleSize(
  population: number,
  riskLevel: string,
  frequency: string
): { sampleSize: number; error: string } {
  const expectedPopValues = [250, 52, 24, 12, 4, 1];

  if (frequency === "As Needed") {
    // Check if population matches a standard expected population value
    if (expectedPopValues.includes(population)) {
      return {
        sampleSize: 0,
        error: `Error: Population ${population} matches a standard expected population value. Please select a specific frequency instead of As Needed.`
      };
    }

    if (population > 250) {
      const sizeMap: Record<string, number> = { "Low": 30, "Medium": 45, "High": 60 };
      return { sampleSize: sizeMap[riskLevel], error: "" };
    } else {
      return { sampleSize: interpolateSampleSize(population, riskLevel), error: "" };
    }
  } else {
    // Handle specific frequency
    const expectedPop = getExpectedPop(frequency);

    if (frequency === "Multiple per day") {
      if (population <= 250) {
        return {
          sampleSize: 0,
          error: `Error: 'Multiple per day' requires population > 250. Current: ${population}`
        };
      }
    } else {
      if (population !== expectedPop) {
        return {
          sampleSize: 0,
          error: `Error: Population ${population} is not valid for frequency '${frequency}'. Expected population: ${expectedPop}`
        };
      }
    }

    return { sampleSize: getSampleSizeFromTable(frequency, riskLevel), error: "" };
  }
}

function getExpectedPop(frequency: string): number {
  const map: Record<string, number> = {
    "Multiple per day": 250,
    "Daily": 250,
    "Weekly": 52,
    "Bi-Weekly": 24,
    "Monthly": 12,
    "Quarterly": 4,
    "Annually": 1
  };
  return map[frequency] || 0;
}

function getSampleSizeFromTable(frequency: string, riskLevel: string): number {
  const table: Record<string, Record<string, number>> = {
    "Multiple per day": { "Low": 30, "Medium": 45, "High": 60 },
    "Daily":           { "Low": 20, "Medium": 25, "High": 30 },
    "Weekly":          { "Low": 5,  "Medium": 8,  "High": 10 },
    "Bi-Weekly":       { "Low": 3,  "Medium": 6,  "High": 8 },
    "Monthly":         { "Low": 2,  "Medium": 3,  "High": 5 },
    "Quarterly":       { "Low": 1,  "Medium": 2,  "High": 2 },
    "Annually":        { "Low": 1,  "Medium": 1,  "High": 1 }
  };
  return table[frequency]?.[riskLevel] || 0;
}

function interpolateSampleSize(population: number, riskLevel: string): number {
  const tables: Record<string, number[][]> = {
    "Low":    [[250, 20], [52, 5], [24, 3], [12, 2], [4, 1], [1, 1]],
    "Medium": [[250, 25], [52, 8], [24, 6], [12, 3], [4, 2], [1, 1]],
    "High":   [[250, 30], [52, 10], [24, 8], [12, 5], [4, 2], [1, 1]]
  };
  const table = tables[riskLevel];

  // Check for exact match
  for (const [pop, size] of table) {
    if (population === pop) return size;
  }

  // Find two points to interpolate between
  for (let i = 0; i < table.length - 1; i++) {
    const [x1, y1] = table[i];
    const [x2, y2] = table[i + 1];

    if (population < x1 && population > x2) {
      const result = y1 + ((population - x1) * (y2 - y1)) / (x2 - x1);
      return Math.ceil(result); // Round up for conservative approach
    }
  }

  // If population is larger than largest in table
  if (population > table[0][0]) return table[0][1];

  // If population is smaller than smallest in table
  if (population < table[table.length - 1][0]) return table[table.length - 1][1];

  return table[0][1];
}
