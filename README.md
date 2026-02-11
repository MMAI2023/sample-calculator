# Sample Calculator Repository

This repository contains audit sample size calculator implementations in multiple formats, designed to help auditors determine appropriate sample sizes based on population, risk level, and frequency.

## 📁 Repository Contents

- **sample1** - VBA/Excel implementation (Base Sample Calculator)
- **sample5** - Extended VBA implementation
- **sample6** - Additional VBA variant
- **sample8** - C implementation
- **Sample7** - TypeScript/Office Scripts implementation for Power Automate

## 🚀 Quick Start

### For Excel Users (VBA)
1. Open the `sample1` file to view the VBA code
2. Follow the setup instructions at the top of the file
3. Run the `SetupCalculator` macro to create the calculator interface

### For Power Automate Users
1. Use the `Sample7` file (TypeScript Office Script)
2. See **[POWER_AUTOMATE_MONDAY_INTEGRATION.md](./POWER_AUTOMATE_MONDAY_INTEGRATION.md)** for integration guide

## 📖 Documentation

- **[Quick Start Guide](./QUICKSTART.md)** - Get started in 20 minutes
- **[Power Automate + Monday.com Integration Guide](./POWER_AUTOMATE_MONDAY_INTEGRATION.md)** - Complete guide on connecting Power Automate with Monday.com to automate sample size calculations

## 💡 Use Cases

- **Audit Planning**: Calculate required sample sizes for audits
- **Quality Control**: Determine sampling requirements for QC processes
- **Risk Assessment**: Adjust sample sizes based on risk levels
- **Automated Workflows**: Integrate with Monday.com via Power Automate

## 🔧 Features

### Base Calculator (sample1, sample5, sample6)
- Excel-based VBA implementation
- Interactive worksheet interface
- Support for various frequencies (Daily, Weekly, Monthly, etc.)
- Risk-based sample size adjustment (Low, Medium, High)
- Sub-population calculations

### Power Automate Integration (Sample7)
- TypeScript Office Scripts
- RESTful API integration ready
- Compatible with Excel Online
- Supports Monday.com workflows
- Handles both base and sub-population calculations

### C Implementation (sample8)
- Standalone C implementation
- Command-line interface
- Portable across platforms

## 📋 Sample Size Calculation

The calculator determines sample sizes based on three key inputs:

1. **Population**: The total number of items/transactions
2. **Risk Level**: Low, Medium, or High
3. **Frequency**: How often the activity occurs
   - Multiple per day (>250 population)
   - Daily (250 population)
   - Weekly (52 population)
   - Bi-Weekly (24 population)
   - Monthly (12 population)
   - Quarterly (4 population)
   - Annually (1 population)
   - As Needed (custom populations)

## 🔗 Integration Options

### Power Automate + Monday.com
Connect your Monday.com boards to automatically calculate sample sizes. See the [integration guide](./POWER_AUTOMATE_MONDAY_INTEGRATION.md) for details.

**Benefits**:
- Automated calculations when items are created
- Real-time updates to Monday.com
- Reduced manual errors
- Complete audit trail

### Other Integrations
The TypeScript implementation (Sample7) can be adapted for:
- SharePoint workflows
- Teams bots
- Custom web applications
- API endpoints

## 🛠️ Technical Details

### VBA Implementation
- Compatible with Excel 2016 and later
- No external dependencies required
- Creates dedicated worksheet interface
- Includes input validation

### TypeScript Implementation
- Requires Office Scripts (Excel Online)
- Uses Excel Script API
- Returns structured JSON results
- Error handling included

### C Implementation
- Standard C library only
- Compiles with GCC/Clang
- Cross-platform compatible

## 📝 License

Please check with the repository owner for licensing information.

## 🤝 Contributing

This is a sample calculator implementation. For questions or issues, please open an issue in the repository.

## 📞 Support

For specific integration support:
- **Power Automate + Monday.com**: See [POWER_AUTOMATE_MONDAY_INTEGRATION.md](./POWER_AUTOMATE_MONDAY_INTEGRATION.md)
- **VBA Issues**: Review comments in the sample1 file
- **Office Scripts**: Review Sample7 implementation

---

**Note**: This calculator provides sample size recommendations based on standard auditing practices. Always consult with qualified auditing professionals for your specific requirements.
