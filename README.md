## xlTemplate

[Windows](#requirements) | [Excel 365](#requirements) | [Excel 2024](#requirements) | [Excel 2021](#requirements) | [License](LICENSE)

A native Microsoft Excel XLL add-in for generating dynamic text and SQL from templates and Excel values.

xlTemplate is a sister project of [DuckDBExcelAddin](https://github.com/your-account/DuckDBExcelAddin). It is designed to generate SQL that can then be passed to the `DUCKDB.EXEC` family of worksheet functions.

### Quick Start

```excel
=RENDER(
    "SELECT * FROM {{ table_name }} WHERE amount > {{ minimum }}",
    "table_name", "sales",
    "minimum", 1000
)
```

Result:

```sql
SELECT * FROM sales WHERE amount > 1000
```

Use the generated SQL with DuckDBExcelAddin:

```excel
=DUCKDB.EXEC(
    RENDER(
        "SELECT * FROM {{ table_name }} WHERE amount > ?",
        "table_name", "sales"
    ),
    A1:D100,
    1000
)
```

### Why SQL Templating?

DuckDBExcelAddin supports parameter binding and parameter binding should remain the preferred method for data values:

```excel
=DUCKDB.EXEC(
    "SELECT * FROM xlrange(?) WHERE cif = ?",
    A1:D100,
    1,
    10001
)
```

However, bound parameters cannot replace every SQL grammar element. In particular, parameters generally cannot be used where SQL requires an identifier, keyword, clause, or generated statement structure.

Templates can generate those structural parts before the resulting SQL is sent to DuckDBExcelAddin.

Examples include:

- Database, schema, table, and column identifiers
- `ATTACH` aliases
- `COPY ... TO` destinations and options
- Generated `CASE` expressions
- Generated projection lists
- Generated filters and grouping clauses
- Dynamic pivot value lists
- Repeated SQL statements

Parameter binding and SQL templating serve different purposes:

- Use **parameter binding** for values whenever possible.
- Use **templating** for identifiers and SQL structure that cannot be bound.

> [!IMPORTANT]
> Template values are inserted into the generated text. They are not automatically SQL-escaped or validated. Only use trusted values for raw SQL fragments and validate dynamic identifiers before execution.

### Features

- Native Excel XLL implementation
- Template rendering directly from worksheet formulas
- Named template values supplied as Excel key/value pairs
- Integer, floating-point, Boolean, and string values
- UTF-8 template rendering
- Dynamic output-buffer growth
- Dynamic template-program growth
- No .NET runtime required
- Designed to complement DuckDBExcelAddin

### Status

**Experimental**

The project is under active development. Formula names, template behavior, supported value types, and error handling may change before the first stable release.

## Installation

### Requirements

- Microsoft Excel 64-bit with Dynamic Array support:
  - Microsoft 365 Excel
  - Excel 2024
  - Excel 2021
- Windows 64-bit

### Enable the Add-in

Copy the following file to a local folder:

```text
xlTemplate.xll
```

Open the XLL directly, or add it through the Excel Add-ins dialog.

## Tutorial

### Basic Replacement

```excel
=RENDER(
    "Hello, {{ name }}!",
    "name", "Excel"
)
```

Result:

```text
Hello, Excel!
```

### Numeric Values

```excel
=RENDER(
    "Quantity: {{ quantity }}, Price: {{ price }}",
    "quantity", 10,
    "price", 12.5
)
```

### Boolean Values

Boolean values are rendered as `TRUE` or `FALSE`:

```excel
=RENDER(
    "Enabled: {{ enabled }}",
    "enabled", TRUE
)
```

### Use with DuckDBExcelAddin

A common workflow is:

1. Use `RENDER` to generate the SQL structure.
2. Pass the generated SQL to `DUCKDB.EXEC` or another DuckDBExcelAddin execution function.
3. Continue using parameter binding for ordinary data values.

```excel
=LET(
    sql,
    RENDER(
        "SELECT {{ columns }} FROM xlrange(?) WHERE amount >= ?",
        "columns", "customer_id, SUM(amount) AS total"
    ),
    DUCKDB.EXEC(sql, A1:D1000, 1, F1)
)
```

### Dynamic `ATTACH`

A database alias is part of SQL syntax and cannot normally be supplied as a bound value:

```excel
=RENDER(
    "ATTACH '{{ path }}' AS {{ alias }}; SELECT * FROM {{ alias }}.main.sales;",
    "path", "C:\\data\\sales.duckdb",
    "alias", "sales_db"
)
```

Generated SQL:

```sql
ATTACH 'C:\data\sales.duckdb' AS sales_db;
SELECT * FROM sales_db.main.sales;
```

### Dynamic `COPY ... TO`

```excel
=RENDER(
    "COPY (SELECT * FROM {{ table_name }}) TO '{{ output_path }}' (FORMAT {{ format }});",
    "table_name", "sales",
    "output_path", "C:\\exports\\sales.parquet",
    "format", "PARQUET"
)
```

Generated SQL:

```sql
COPY (SELECT * FROM sales)
TO 'C:\exports\sales.parquet'
(FORMAT PARQUET);
```

### Generate a `CASE` Expression

Templates can be used to select or generate SQL expressions based on worksheet values:

```excel
=RENDER(
    "SELECT CASE {{ mode }} WHEN 1 THEN amount WHEN 2 THEN quantity ELSE 0 END AS result FROM sales",
    "mode", 2
)
```

For more complex generated branches, build the branch text in Excel and insert it as a trusted SQL fragment:

```excel
=LET(
    case_clause,
    SWITCH(
        A1,
        "amount", "CASE WHEN amount >= 1000 THEN 'High' ELSE 'Low' END",
        "status", "CASE WHEN active THEN 'Active' ELSE 'Inactive' END",
        "NULL"
    ),
    RENDER(
        "SELECT customer_id, {{ case_clause }} AS category FROM sales",
        "case_clause", case_clause
    )
)
```

### Combine Templating and Parameter Binding

Do not insert ordinary data values into SQL text when they can be bound safely.

Preferred:

```excel
=LET(
    sql,
    RENDER(
        "SELECT * FROM {{ table_name }} WHERE customer_id = ?",
        "table_name", "sales"
    ),
    DUCKDB.EXEC(sql, A1:D1000, 10001)
)
```

Avoid inserting an untrusted value directly into SQL:

```excel
=RENDER(
    "SELECT * FROM sales WHERE customer_id = {{ customer_id }}",
    "customer_id", A1
)
```

## Formula Reference

### `RENDER`

```text
=RENDER(template, [key1], [value1], [key2], [value2], ...)
```

#### Parameters

- `template`: Required template text.
- `key1`, `key2`, ...: Template variable names.
- `value1`, `value2`, ...: Values associated with the preceding keys.

The current implementation supports up to 64 key/value pairs.

Supported value types:

- Integer
- Number
- Boolean
- String

Keys must be strings and must be unique. Key/value pairs must be continuous. A missing pair between supplied pairs is treated as invalid input.

### `RENDER.INFO`

```text
=RENDER.INFO()
```

Returns xlTemplate add-in version information.

## Template Syntax

xlTemplate uses the TinyTemplate engine.

### Variable Output

```text
{{ variable_name }}
```

Example:

```excel
=RENDER("SELECT * FROM {{ table_name }}", "table_name", "sales")
```

### Expressions and Control Flow

Template syntax support follows the bundled TinyTemplate version. Refer to the TinyTemplate project documentation and tests for the exact supported expression, conditional, and iteration syntax.

## Architecture

xlTemplate runs inside the Excel process as a native XLL add-in.

The rendering flow is:

1. Convert the Excel template string from Excel UTF-16 format to UTF-8.
2. Compile the UTF-8 template into a TinyTemplate instruction program.
3. Grow the instruction buffer automatically when the initial capacity is insufficient.
4. Parse Excel key/value arguments into a UTF-8 lookup dictionary.
5. Evaluate the compiled template.
6. Append writer callbacks into a dynamically growing UTF-8 output buffer.
7. Convert the final UTF-8 result to an Excel string.
8. Return the string to Excel.

String parameters are converted to UTF-8 once during parameter parsing and remain owned by the rendering context until evaluation finishes.

## Relationship with DuckDBExcelAddin

xlTemplate does not execute SQL itself. It generates text.

DuckDBExcelAddin executes SQL and returns DuckDB query results to Excel. The two projects are intended to work together:

```text
Excel values
    -> xlTemplate RENDER
    -> generated SQL
    -> DuckDBExcelAddin DUCKDB.EXEC
    -> DuckDB result
    -> Excel dynamic array
```

This separation keeps the responsibilities clear:

- xlTemplate handles dynamic SQL construction.
- DuckDBExcelAddin handles ranges, parameter binding, execution, and result conversion.

## Security Guidance

SQL templating performs text generation, not parameter binding.

Values inserted into SQL may change the meaning of the generated statement. A malicious or malformed value may produce invalid SQL or SQL injection.

Recommended practices:

- Use DuckDB parameter binding for data values.
- Use templates only where binding cannot represent the required SQL structure.
- Accept dynamic identifiers from controlled lists rather than arbitrary text.
- Validate database aliases, table names, column names, formats, and clause fragments.
- Quote and escape string literals correctly when they must be generated.
- Review the rendered SQL before executing templates sourced from users or external workbooks.

## Known Limitations

### Templates Do Not Escape SQL

xlTemplate is a general text-template engine. It does not know whether a value represents:

- A SQL identifier
- A SQL string literal
- A keyword
- A path
- A complete SQL fragment

The caller is responsible for validation, quoting, and escaping.

### Excel String Length

Rendered text must fit within Excel's string-length limit. Results exceeding the Excel limit are not supported.

### Empty or Invalid Results

The add-in may return `#VALUE!` when:

- Template conversion fails
- Template compilation fails
- A key is missing or duplicated
- A key or value type is unsupported
- Template evaluation fails
- Memory allocation fails
- The rendered output cannot be converted to an Excel string

### Parameter Keys

- Keys must be strings.
- Keys must be unique.
- A value must follow every supplied key.
- Missing key/value pairs are allowed only after the final supplied pair.

### SQL Execution

xlTemplate only renders text. DuckDBExcelAddin or another SQL client is required to execute generated SQL.

## Troubleshooting

### Add-in Fails to Load

Verify that:

- Excel is 64-bit.
- The XLL is not blocked by Windows.
- The add-in was built for the same architecture as Excel.
- Required runtime libraries are available.

### `#VALUE!` Returned

Verify that:

- The template syntax is valid.
- Every referenced variable has a matching key/value pair.
- Keys are strings and are not duplicated.
- Values use a supported Excel type.
- The rendered result does not exceed Excel's string limit.

### Generated SQL Fails

Inspect the generated SQL separately before passing it to DuckDBExcelAddin:

```excel
=RENDER(template_cell, "table_name", table_name_cell)
```

Check identifier quoting, string escaping, paths, commas, parentheses, and clause ordering.

## Building

### Build Requirements

- Excel XLL SDK
- [uthash](https://troydhanson.github.io/uthash/)
- TinyTemplate
- A Windows C/C++ toolchain compatible with the Excel XLL SDK

### Toolchain

Development is primarily intended for MinGW-w64-based environments such as w64devkit. Other toolchains may work but may require Makefile or SDK adjustments.

### Excel SDK Header Note

The Excel SDK `XLCALL.H` header may define an XLOPER member named `bool`, which conflicts with the C `bool` keyword. If required by the selected toolchain, rename that member locally, for example to `xbool`, and update matching references in the SDK framework source.

This local header change does not alter the Excel ABI because the member name is not part of the binary layout.

### Build Instructions

Development build:

```shell
make EXCEL_SDK_PATH=<excel-sdk> xll
```

Release build:

```shell
make ADDIN_VERSION=vx.x.x EXCEL_SDK_PATH=<excel-sdk> xll
```

Use Clang instead of the default compiler:

```shell
make CC=clang CXX=clang++ EXCEL_SDK_PATH=<excel-sdk> xll
```

### Output

```text
dist/xlTemplate.xll
```

## Roadmap

Potential future work includes:

- SQL-oriented filters for quoting identifiers and literals
- Better template compilation and evaluation error messages
- Expanded examples for DuckDBExcelAddin integration
- Additional template value types
- Workbook-based regression tests
- Reusable SQL-template libraries

## Acknowledgements

Special thanks to the TinyTemplate contributors for the template engine used by this project.

Additional thanks to:

- [uthash](https://troydhanson.github.io/uthash/)
- Microsoft Excel XLL SDK
- DuckDB and DuckDBExcelAddin contributors

## License

MIT License
