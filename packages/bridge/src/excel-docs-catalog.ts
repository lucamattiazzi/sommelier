/** Curated summaries and synthetic examples; source pages are linked, never fetched at runtime. */
export type ExcelDoc = {
  readonly id: string;
  readonly title: string;
  readonly aliases: string[];
  readonly summary: string;
  readonly syntax: string;
  readonly examples: { readonly formula: string; readonly explanation: string }[];
  readonly notes: string[];
  readonly compatibility: string;
  readonly sources: string[];
};

export const excelDocCatalog: readonly ExcelDoc[] = [
  {
    id: "sum",
    title: "SUM",
    aliases: ["somma", "total", "totale"],
    summary: "Add numbers in a bounded range.",
    syntax: "SUM(number1, [number2], ...)",
    examples: [
      {
        formula: "=SUM(C2:C20)",
        explanation: "Total the numeric amounts in C2:C20.",
      },
    ],
    notes: [
      "Text in referenced cells is ignored; inspect numbers stored as text if totals are unexpected.",
    ],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/sum-function"],
  },
  {
    id: "sumif",
    title: "SUMIF",
    aliases: ["somma.se", "conditional sum", "somma condizione"],
    summary: "Sum values matching one criterion.",
    syntax: "SUMIF(range, criteria, [sum_range])",
    examples: [
      {
        formula: '=SUMIF(A2:A20,"North",C2:C20)',
        explanation: "Total column C for rows whose region in column A is North.",
      },
    ],
    notes: [
      "SUMIF puts the criteria range first; SUMIFS puts the sum range first.",
      'For a threshold from a cell, concatenate the operator: ">="&E2.',
    ],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/sumif-function"],
  },
  {
    id: "sumifs",
    title: "SUMIFS",
    aliases: ["somma.più.se", "multiple criteria", "somma criteri"],
    summary: "Sum values only where every criterion matches.",
    syntax: "SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
    examples: [
      {
        formula: '=SUMIFS(C2:C20,A2:A20,"North",B2:B20,">=5")',
        explanation: "Total amounts for North rows with a quantity of at least five.",
      },
    ],
    notes: [
      "Criteria ranges must have the same dimensions as sum_range.",
      "Conditions are combined with AND; summing overlapping OR groups can double count.",
    ],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/sumifs-function"],
  },
  {
    id: "countif",
    title: "COUNTIF",
    aliases: ["conta.se", "count matches", "conteggio"],
    summary: "Count cells that match one criterion.",
    syntax: "COUNTIF(range, criteria)",
    examples: [
      {
        formula: '=COUNTIF(B2:B20,">=5")',
        explanation: "Count entries of at least five.",
      },
    ],
    notes: ["Use COUNTIFS for multiple criteria. Wildcards * and ? match text; ~ escapes them."],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: ["https://support.microsoft.com/en-us/excel/count-how-often-a-value-occurs-in-excel"],
  },
  {
    id: "countifs",
    title: "COUNTIFS",
    aliases: ["conta.più.se", "count multiple criteria"],
    summary: "Count rows meeting all supplied conditions.",
    syntax: "COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
    examples: [
      {
        formula: '=COUNTIFS(A2:A20,"North",B2:B20,">=5")',
        explanation: "Count North rows whose quantity is at least five.",
      },
    ],
    notes: ["All criteria ranges must have matching dimensions. Conditions use AND."],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/countifs-function"],
  },
  {
    id: "if",
    title: "IF",
    aliases: ["se", "condition", "condizione"],
    summary: "Choose between two results based on a logical test.",
    syntax: "IF(logical_test, value_if_true, [value_if_false])",
    examples: [
      {
        formula: '=IF(B2>=5,"Ready","Pending")',
        explanation: "Classify a quantity using a threshold.",
      },
    ],
    notes: ["Quote literal text, but do not quote cell references."],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/if-function"],
  },
  {
    id: "iferror",
    title: "IFERROR",
    aliases: ["se.errore", "error handling", "gestione errori"],
    summary: "Return a fallback when an expression produces an Excel error.",
    syntax: "IFERROR(value, value_if_error)",
    examples: [
      {
        formula: '=IFERROR(C2/B2,"Check quantity")',
        explanation: "Display a diagnostic message if the division fails.",
      },
    ],
    notes: [
      "Do not hide unexpected errors with zero or empty text before investigating them.",
      "Use IFNA instead when only a missing lookup should be handled.",
    ],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/iferror-function"],
  },
  {
    id: "xlookup",
    title: "XLOOKUP",
    aliases: ["cerca.x", "lookup", "ricerca", "search key"],
    summary: "Find a key and return the corresponding value; exact matching is the default.",
    syntax:
      "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
    examples: [
      {
        formula: '=XLOOKUP(E2,A2:A20,C2:C20,"Not found",0)',
        explanation: "Find the key from E2 in column A and return its amount from column C.",
      },
    ],
    notes: [
      "Duplicates normally return the first match. Choose a duplicate policy explicitly.",
      "Binary search modes require correctly sorted lookup arrays.",
    ],
    compatibility:
      "Excel 2021/2024, Microsoft 365 and Excel for the web; not available in Excel 2016 or 2019.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/xlookup-function"],
  },
  {
    id: "index-match",
    title: "INDEX and MATCH",
    aliases: ["indice", "confronta", "legacy lookup"],
    summary: "Combine an exact position lookup with retrieval from another range.",
    syntax: "INDEX(return_range, MATCH(key, lookup_range, 0))",
    examples: [
      {
        formula: "=INDEX(C2:C20,MATCH(E2,A2:A20,0))",
        explanation: "Look up E2 without requiring XLOOKUP.",
      },
    ],
    notes: ["MATCH with 0 requests exact matching; omitted match_type has different semantics."],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: [
      "https://support.microsoft.com/en-us/excel/functions/index-function",
      "https://support.microsoft.com/en-us/excel/functions/match-function",
    ],
  },
  {
    id: "filter",
    title: "FILTER",
    aliases: ["filtro", "filter rows", "filtrare righe"],
    summary: "Return the rows or columns selected by a Boolean array.",
    syntax: "FILTER(array, include, [if_empty])",
    examples: [
      {
        formula: '=FILTER(A2:C20,B2:B20>=5,"No matches")',
        explanation: "Return three columns for rows whose quantity is at least five.",
      },
    ],
    notes: [
      "The include array must match the relevant height or width of array.",
      "The result spills into adjacent cells; keep that area free.",
    ],
    compatibility:
      "Excel 2021/2024, Microsoft 365 and Excel for the web; not available in Excel 2016 or 2019.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/filter-function"],
  },
  {
    id: "unique",
    title: "UNIQUE",
    aliases: ["unici", "distinct", "duplicati"],
    summary: "Return distinct rows or columns from an array.",
    syntax: "UNIQUE(array, [by_col], [exactly_once])",
    examples: [
      {
        formula: "=UNIQUE(A2:A20)",
        explanation: "List distinct region values.",
      },
    ],
    notes: [
      "exactly_once=TRUE returns only entries occurring once, rather than all distinct entries.",
      "Inspect blanks in the source if the result contains unexpected empty or zero entries.",
    ],
    compatibility:
      "Excel 2021/2024, Microsoft 365 and Excel for the web; not available in Excel 2016 or 2019.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/unique-function"],
  },
  {
    id: "sort",
    title: "SORT",
    aliases: ["dati.ordina", "sorting", "ordinamento"],
    summary: "Sort an array by a chosen row or column.",
    syntax: "SORT(array, [sort_index], [sort_order], [by_col])",
    examples: [
      {
        formula: "=SORT(A2:C20,3,-1)",
        explanation: "Order three-column rows by column C, descending.",
      },
    ],
    notes: [
      "sort_index is relative to the input array, starting at 1.",
      "The result spills and does not reorder the source cells.",
    ],
    compatibility:
      "Excel 2021/2024, Microsoft 365 and Excel for the web; not available in Excel 2016 or 2019.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/sort-function"],
  },
  {
    id: "let",
    title: "LET",
    aliases: ["named expression", "variabili", "readability"],
    summary: "Name intermediate calculations within one formula.",
    syntax: "LET(name1, value1, [name2, value2], ..., calculation)",
    examples: [
      {
        formula: "=LET(amount,C2,rate,D2,amount*(1-rate))",
        explanation: "Name the amount and discount rate before computing the net amount.",
      },
    ],
    notes: [
      "The last argument is the expression to return.",
      "Names must obey Excel naming rules; avoid names that conflict with cell or R1C1 references.",
    ],
    compatibility:
      "Excel 2021/2024, Microsoft 365 and Excel for the web; not available in Excel 2016 or 2019.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/let-function"],
  },
  {
    id: "lambda",
    title: "LAMBDA",
    aliases: ["custom function", "funzione personalizzata"],
    summary: "Define a reusable calculation with explicit parameters.",
    syntax: "LAMBDA([parameter1, parameter2, ...], calculation)",
    examples: [
      {
        formula: "=LAMBDA(amount,rate,amount*(1-rate))(120,0.15)",
        explanation: "Define and immediately call a net-amount calculation.",
      },
    ],
    notes: [
      "An uncalled LAMBDA in a cell returns #CALC!.",
      "Workbook-level reuse needs a defined name; Sommelier does not currently expose a tool for creating defined names.",
    ],
    compatibility:
      "Excel 2024, Microsoft 365 and Excel for the web; do not assume availability in Excel 2021 or earlier.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/lambda-function"],
  },
  {
    id: "date",
    title: "DATE",
    aliases: ["data", "dates", "date serial"],
    summary: "Construct a date from numeric year, month and day components.",
    syntax: "DATE(year, month, day)",
    examples: [
      {
        formula: "=DATE(2026,9,22)",
        explanation: "Build an unambiguous date without parsing a regional date string.",
      },
    ],
    notes: [
      "Excel represents dates using serial numbers; display formatting is separate.",
      "DATE normalizes out-of-range month/day components, so it is not an input validator.",
    ],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: ["https://support.microsoft.com/en-us/excel/functions/date-function"],
  },
  {
    id: "references",
    title: "Relative, absolute and mixed references",
    aliases: [
      "riferimenti assoluti",
      "riferimenti relativi",
      "absolute references",
      "relative references",
      "mixed references",
      "dollar",
    ],
    summary: "Choose which row and column coordinates remain fixed when a formula is copied.",
    syntax: "A1; $A$1; $A1; A$1",
    examples: [
      {
        formula: "=B2*$F$1",
        explanation: "Multiply a changing row value by a fixed rate in F1.",
      },
    ],
    notes: [
      "Dollar signs anchor a row or column for copying; they do not prevent edits.",
      "Quote sheet names containing spaces, for example ='Sales Data'!B2.",
    ],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: [
      "https://support.microsoft.com/en-us/excel/switch-between-relative-absolute-and-mixed-references",
    ],
  },
  {
    id: "tables",
    title: "Structured references in Excel tables",
    aliases: ["tabelle", "riferimenti strutturati", "structured references", "table"],
    summary: "Refer to table columns and current-row values by name.",
    syntax: "TableName[Column]; [@Column]",
    examples: [
      {
        formula: "=SUM(Sales[Amount])",
        explanation: "Total the Amount column in an existing table named Sales.",
      },
    ],
    notes: [
      "Discover actual table and column names before composing the formula.",
      "[@Amount] refers to the current row in a table; Sales[Amount] refers to its data column.",
      "Creating or renaming tables is separate from writing a formula.",
    ],
    compatibility:
      "Excel 2016 and later, Microsoft 365 and Excel for the web. Check the linked documentation for your host.",
    sources: [
      "https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables",
    ],
  },
  {
    id: "dynamic-arrays",
    title: "Dynamic arrays and spill errors",
    aliases: ["spill", "espansione", "matrici dinamiche", "dynamic arrays"],
    summary: "Understand formulas whose results occupy multiple adjacent cells.",
    syntax: "anchor#; @expression",
    examples: [
      {
        formula: "=SUM(F2#)",
        explanation: "Sum the complete spilled result starting at F2.",
      },
    ],
    notes: [
      "#SPILL! can indicate occupied cells or an unsupported placement; never clear neighboring data without approval.",
      "Place spilled formulas outside Excel tables.",
      "@ requests implicit intersection and can change the meaning of an array formula.",
    ],
    compatibility:
      "Excel 2021/2024, Microsoft 365 and Excel for the web; not available in Excel 2016 or 2019.",
    sources: [
      "https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior",
    ],
  },
  {
    id: "locale",
    title: "Formula language and regional settings",
    aliases: ["italiano", "english", "localization", "separatori", "formulaslocal", "locale"],
    summary:
      "Distinguish canonical formulas from formulas entered using local language and separators.",
    syntax: "Office.js Range.formulas versus Range.formulasLocal",
    examples: [
      {
        formula: "=SUM(A2:A5)",
        explanation:
          "Canonical English example; localized names and separators depend on the API and Excel settings.",
      },
    ],
    notes: [
      "Do not replace commas or translate function names blindly; quoted text and array constants require care.",
      "Range.formulasLocal uses the user's language and formatting conventions.",
      "Sommelier currently exposes range.write as values assignment, not a dedicated formulasLocal API. Read the resulting formula and value back; do not assume automatic translation.",
    ],
    compatibility: "Depends on the Excel host, regional settings and assignment API.",
    sources: [
      "https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.1#excel-excel-range-formulaslocal-member",
    ],
  },
];
