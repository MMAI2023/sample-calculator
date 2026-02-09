/**
 * AUDIT SAMPLING CALCULATOR - Power Automate Office Script
 *
 * This Office Script can be called directly from Power Automate
 * to calculate both base sample size and sub-sample size.
 *
 * FUNCTION: main (Base Sampling - Part 1)
 * INPUTS:
 *   population : number  - The actual population size (minimum 1)
 *   riskLevel  : string  - "Low", "Medium", or "High"
 *   frequency  : string  - "As Needed", "Multiple per day", "Daily", "Weekly",
 *                           "Bi-Weekly", "Monthly", "Quarterly", "Annually"
 *
 * OUTPUTS (returned object):
 *   baseSampleSize : number  - The calculated base sample size (0 if error)
 *   error          : string  - Error message (empty string if successful)
 *
 * FUNCTION: calculateSubSampleResult (Sub-Sampling - Part 2)
 * INPUTS:
 *   subPopulation : number  - The sub-population size (minimum 1)
 *   riskLevel     : string  - "Low", "Medium", or "High" (same as Part 1)
 *   frequency     : string  - "As Needed"
 *
 * OUTPUTS (returned object):
 *   subSampleSize : number  - The calculated sub-sample size (0 if error)
 *   error         : string  - Error message (empty string if successful)
 */
function main(
  workbook: ExcelScript.Workbook,
  population: number,
  riskLevel: string,
  frequency: string
): { baseSampleSize: number; error: string } {
  // Validate inputs
  if (population < 1) {
    return { baseSampleSize: 0, error: "Please enter a valid population (minimum 1)" };
  }
  if (!["Low", "Medium", "High"].includes(riskLevel)) {
    return { baseSampleSize: 0, error: "Please select a valid risk level (Low, Medium, or High)" };
  }
  const validFrequencies = [
    "As Needed", "Multiple per day", "Daily", "Weekly",
    "Bi-Weekly", "Monthly", "Quarterly", "Annually"
  ];
  if (!validFrequencies.includes(frequency)) {
    return { baseSampleSize: 0, error: "Please select a valid frequency" };
  }

  const result = calculateSampleSize(population, riskLevel, frequency);
  return { baseSampleSize: result.sampleSize, error: result.error };
}

/**
 * Call this function from a separate Power Automate "Run script" action
 * for Part 2 sub-sampling calculation.
 */
function calculateSubSampleResult(
  workbook: ExcelScript.Workbook,
  subPopulation: number,
  riskLevel: string,
  frequency: string
): { subSampleSize: number; error: string } {
  // Validate inputs
  if (subPopulation < 1) {
    return { subSampleSize: 0, error: "Please enter a valid sub-population (minimum 1)" };
  }
  if (!["Low", "Medium", "High"].includes(riskLevel)) {
    return { subSampleSize: 0, error: "Please complete Part 1 first (Risk Level required)" };
  }
  if (frequency !== "As Needed") {
    return { subSampleSize: 0, error: "Sub-sampling frequency must be 'As Needed'" };
  }

  const result = calculateSampleSize(subPopulation, riskLevel, frequency);
  return { subSampleSize: result.sampleSize, error: result.error };
}

function calculateSampleSize(
  population: number,
  riskLevel: string,
  frequency: string
): { sampleSize: number; error: string } {
  const expectedPopValues = [250, 52, 24, 12, 4, 1];

  if (frequency === "As Needed") {
    if (expectedPopValues.includes(population)) {
      return {
        sampleSize: 0,
        error: `Error: Population ${population} matches standard value. Select specific frequency.`
      };
    }

    if (population > 250) {
      const sizeMap: Record<string, number> = { "Low": 30, "Medium": 45, "High": 60 };
      return { sampleSize: sizeMap[riskLevel] || 45, error: "" };
    } else {
      return { sampleSize: interpolateSampleSize(population, riskLevel), error: "" };
    }
  } else {
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
          error: `Error: Population ${population} invalid for ${frequency}. Expected: ${expectedPop}`
        };
      }
    }

    return { sampleSize: getSampleSizeFromTable(frequency, riskLevel), error: "" };
  }
}

function getExpectedPop(frequency: string): number {
  const map: Record<string, number> = {
    "Multiple per day": 251,
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
    "Quarterly":       { "Low": 2,  "Medium": 2,  "High": 2 },
    "Annually":        { "Low": 1,  "Medium": 1,  "High": 1 }
  };
  return table[frequency]?.[riskLevel] || 0;
}

function interpolateSampleSize(population: number, riskLevel: string): number {
  const tables: Record<string, number[][]> = {
    "Low":    [[250, 20], [52, 5], [24, 3], [12, 2], [4, 2], [1, 1]],
    "Medium": [[250, 25], [52, 8], [24, 6], [12, 3], [4, 2], [1, 1]],
    "High":   [[250, 30], [52, 10], [24, 8], [12, 5], [4, 2], [1, 1]]
  };
  const table = tables[riskLevel] || tables["Medium"];

  for (const [pop, size] of table) {
    if (population === pop) return size;
  }

  for (let i = 0; i < table.length - 1; i++) {
    const [x1, y1] = table[i];
    const [x2, y2] = table[i + 1];

    if (population < x1 && population > x2) {
      const result = y1 + ((population - x1) * (y2 - y1)) / (x2 - x1);
      return Math.ceil(result);
    }
  }

  if (population > table[0][0]) return table[0][1];
  if (population < table[table.length - 1][0]) return table[table.length - 1][1];

  return table[0][1];
}
