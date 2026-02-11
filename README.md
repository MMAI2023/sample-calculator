# Sample Calculator

A collection of audit sampling calculators implemented across multiple platforms — VBA, Office Scripts, and Power Automate (TypeScript).

## Samples

| Directory | Platform | Description |
|-----------|----------|-------------|
| `sample1` | VBA Macro | Base sample calculator |
| `sample5` | VBA Macro | Enhanced version with Part 2 (sub-sampling) |
| `sample6` | VBA Macro | Worksheet event handler — auto-clears results on input change |
| `Sample7` | TypeScript (Power Automate) | Power Automate entrypoint for Excel Script runtime |
| `sample8` | Office Scripts | Excel Online compatible calculator with logging |

## Features

- Three risk levels: **Low**, **Medium**, **High**
- Eight frequency options: As Needed, Multiple per day, Daily, Weekly, Bi-Weekly, Monthly, Quarterly, Annually
- Linear interpolation for custom population sizes
- Input validation and error messages
- Two-part calculator (base + sub-sampling)

## Documentation

- [Connecting Power Automate to Monday.com](docs/power-automate-monday-connection.md) — Step-by-step guide and simple analysis workflow
