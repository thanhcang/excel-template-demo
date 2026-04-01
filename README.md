# Excel COA Export with TypeScript

A TypeScript project that uses `xlsx-template` to populate a Certificate of Analysis (COA) Excel template with data.

## Features

- ✓ Uses `xlsx-template` library for Excel template management
- ✓ Loads COA.xlsx template and substitutes placeholders with data
- ✓ Generates COA_Output.xlsx with populated information
- ✓ TypeScript support with strict type checking
- ✓ Easy to extend with additional data fields

## Installation

```bash
npm install
```

## Usage

### Run the project

```bash
npm start
```

Or build and run separately:

```bash
npm run build
node dist/index.js
```

## Data Structure

The project populates the following fields:

```typescript
{
  invoice_no: "FG-DF-0309-01",
  container_no: "FBIU5605008",
  serial_no: "SITR486555",
  date: "March 30, 2026",
  batch_no: "20003/0021",
  production_date: "21/4/2026",
  expired_date: "20/4/28"
}
```

## Template Requirements

The `COA.xlsx` template file must contain placeholder text in the following format:

```
{{invoice_no}}
{{container_no}}
{{serial_no}}
{{date}}
{{batch_no}}
{{production_date}}
{{expired_date}}
```

### How to Create a Template

1. Open Excel and create your COA layout
2. In the cells where you want data to appear, use placeholders like `{{field_name}}`
3. Save the file as `COA.xlsx` in the project root
4. Run the project to generate the populated output

## Output

After running, the generated file `COA_Output.xlsx` will be created with all placeholders replaced by actual data.

## Project Structure

```
.
├── src/
│   ├── index.ts          # Main application
│   └── inspect.ts        # Utility script to inspect template structure
├── COA.xlsx             # Input template
├── COA_Output.xlsx      # Generated output
├── package.json
├── tsconfig.json
└── README.md
```

## Dependencies

- `xlsx-template`: ^1.3.0 - Excel template engine
- `typescript`: ^5.0.0 - TypeScript compiler
- `ts-node`: ^10.9.0 - TypeScript runtime

## Development

### Inspect Template Structure

To inspect the structure of the Excel template:

```bash
npx ts-node src/inspect.ts
```

This will show sheet names and attempts substitution, helping debug template issues.

## License

MIT
