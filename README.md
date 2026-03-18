# Excel Paste Table Component

A React component that parses pasted Excel-style tabular data and renders it as a styled HTML table.

## Included Files

- `src/components/ExcelPasteTable.jsx`
- `src/components/ExcelPasteTable.css`
- `src/App.jsx`
- `src/main.jsx`
- `src/styles.css`

## Features

- Paste area with automatic parse on paste
- Manual `Process Data` parsing button
- `.csv` / `.txt` file upload support
- `Clear` button to reset all inputs and parsed results
- First row mapped to table headers
- Handles empty cells and trims whitespace
- Inconsistent column count normalization (pads missing cells)
- Row/column summary
- Green header theme and row hover highlighting
- Horizontal scrolling for wide tables
- Friendly empty and error states
- Optional: delimiter selection (auto/tab/comma/semicolon)
- Optional: export parsed rows as JSON

## Usage in Existing React App

1. Copy `src/components/ExcelPasteTable.jsx` and `src/components/ExcelPasteTable.css` into your app.
2. Import and render it:

```jsx
import ExcelPasteTable from './components/ExcelPasteTable';

function App() {
  return <ExcelPasteTable />;
}
```

## If Running This Workspace as a React App

If this folder is not already a React project, initialize one first (for example with Vite), then keep the provided `src` files.

```bash
npm create vite@latest . -- --template react
npm install
npm run dev
```

After scaffolding, ensure these files remain:

- `src/main.jsx` should render `App`
- `src/App.jsx` should render `ExcelPasteTable`
