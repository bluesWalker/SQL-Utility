# SQL Utility User Manual

SQL Utility is a portable Windows application for browsing physical SQL Server tables, running approved read-only queries, paging through results, and exporting data to Excel. It connects with your signed-in Windows account and does not ask for or store a database password.

For implementation details, technical constraints, tests, and maintainer information, see the [Technical Reference](TECHNICAL_REFERENCE.md).

## Before we start

Purpose of developing this tool is to provide some convenience for frequent day-to-day queries and outputs processing in databases and it should stay within this scope.

- The app is strictly read-only. No CREATE/UPDATE/DETELE is allowed. For use cases beyond that, please use SQLCMD instead.
- It is designed for internal use only so there is no credential control. It uses only your signed in Windows account.
- It is designed to be lightweight and portable. No third-party dependencies are required.

## Install and start

SQL Utility does not use an installer and does not require administrator access.

1. Copy or extract the complete application folder to a Windows location that you can access. Do not run the application from inside a ZIP file.
2. Keep these eight runtime files together in the same structure:

   ```text
   StartSqlUtility.cmd
   SqlUtility.ps1
   SqlUtility.cat
   modules/
     SqlUtility.Config.ps1
     SqlUtility.QueryPolicy.ps1
     SqlUtility.DataExplorer.ps1
     SqlUtility.Database.ps1
     SqlUtility.Excel.ps1
   ```

3. Double-click `StartSqlUtility.cmd`.

The launcher may show a command window briefly during startup, then hides it while SQL Utility remains open.

Before opening the application, the launcher checks the seven protected command/script files against the SHA-256 hashes in `SqlUtility.cat`. If a protected file is missing or changed, SQL Utility refuses to start and asks you to extract a fresh copy of the original package. `SqlUtility.config.json` and its temporary files are not part of this check because they contain normal saved settings.

Windows PowerShell 5.1 is required. Microsoft Excel is not required to run the application or create an `.xlsx` file, but Excel or another compatible spreadsheet application is needed to open the exported file.

If you want saved connections and settings to remain available after closing the application, place the application in a folder where you have write access. SQL Utility stores those preferences in `SqlUtility.config.json` beside the application. Keep that file with the application when moving it to another folder if you want to retain the preferences.

## Connect to SQL Server

The **Server** and **Database** fields are blank whenever the application starts.

1. Enter the SQL Server name or network alias in **Server**.
2. Enter the database name in **Database**.
3. Select **Test Connection** to check the connection without opening the workspace, or select **Connect** to check it and continue.

The connection uses your signed-in Windows identity. Your account must already have access to the server, database, and data you want to use. The connection attempt stops after 10 seconds if the server does not respond.

After a successful test or connection, the server/database pair appears under **Saved Connections** and is saved when the application folder is writable. Selecting a saved connection fills the two fields but does not connect automatically. Select a saved connection and then **Delete** to remove it; the application asks for confirmation.

After connecting, the workspace shows the active server and database. Select **Change Connection** to return to the connection screen. If you confirm, the current SQL text, query results, Data Explorer selections, filters, and preview are cleared. Saved connections and settings remain.

## Query tab

Use **Query** to enter and run a supported read-only SQL Server `SELECT` statement. The result grid is read-only.

### Write and execute a query

Enter one `SELECT` statement, then select **Execute**. The supported source shape is:

- One named source written as `table` or `schema.table`.
- Optional chained `JOIN`, `INNER JOIN`, `LEFT JOIN`, or `LEFT OUTER JOIN` clauses using named sources and an `ON` condition.
- Optional expressions and clauses such as `DISTINCT`, `WHERE`, `GROUP BY`, `HAVING`, and `ORDER BY` when they remain within the approved single-statement shape.

Example:

```sql
SELECT c.CustomerId, c.Name, o.OrderDate
FROM dbo.Customers AS c
LEFT JOIN dbo.Orders AS o ON o.CustomerId = c.CustomerId
WHERE c.IsActive = 1
ORDER BY c.CustomerId, o.OrderDate;
```

The Query tab rejects data-changing statements and unsupported shapes, including additional statements, subqueries, CTEs, `RIGHT JOIN`, `FULL JOIN`, `CROSS JOIN`, `APPLY`, comma joins, table functions, temporary tables, and cross-database or linked-server sources. This validation helps prevent accidental changes; your SQL Server permissions remain the authoritative access control.

If you edit the SQL after a successful execution, the displayed result becomes stale. Paging, counting, and export are unavailable until you select **Execute** again.

### Page through results

Each displayed page contains up to 500 rows. Use **<** for the previous page and **>** for the next page.

- With a top-level `ORDER BY`, each page request runs the query again for that page. Use a stable, preferably unique ordering. A total count is not calculated automatically.
- Without a top-level `ORDER BY`, SQL Utility retrieves up to the configured **Maximum unordered rows** and pages through that retained data locally. If more rows exist, the status shows the retained number with `+`, export is disabled, and the application asks you to add `ORDER BY`.

Data can change between page requests, so ordered page contents and explicit counts are point-in-time results.

### Count rows

Select **Count** to run a separate count for the current query. Counting is never automatic and may take time for a large or complex query.

- A complete unordered result already has an exact retained count, so **Count** is disabled.
- An ordered or truncated result does not have a total until **Count** succeeds.
- You can select **Count** again later to refresh an explicit total.

### Export a complete result

Select **Export** and choose an `.xlsx` destination. Export is available only when SQL Utility can produce the complete result:

- An ordered query is run again and its complete result is written to the workbook.
- A complete unordered result is exported from the retained data.
- A truncated unordered result cannot be exported. Add a suitable `ORDER BY`, execute again, and then export.

The workbook contains one worksheet named `Results`, with a bold filtered header row that remains visible while scrolling. Existing destination files require confirmation before replacement. The export destination is not remembered.

## Data Explorer tab

Use **Data Explorer** when you know the table you need and want to build a simple query without writing SQL.

### Choose a table

- The tab lists the physical user tables visible to your Windows account. Views are not listed.
- A table in the `dbo` schema appears as its table name. Other tables appear as `schema.table`.
- Type in the box above the table list to filter the displayed names. This filters the list already loaded in the application; it does not query the database again.
- Select **Refresh** to reload the table list from SQL Server.
- Select a table, then select **Preview**. The first preview loads the table's columns, selects every output column, and retrieves data.

Selecting a table by itself does not retrieve its columns or rows.

### Choose output columns

Checked columns are included in the next preview or in SQL sent to the Query tab. At least one column must be checked.

- Select **All** or **None** to check or clear every column.
- Click a checkbox, double-click a column name, or highlight a column and press Space to change its check state.
- A single click on column text only highlights it; it does not change the checkbox.
- While the column list has focus, type the beginning of a column name to jump to it. Characters typed within one second are treated as one prefix.

Changing the checked columns does not alter the preview already displayed. Select **Preview** again when you want to retrieve a new preview.

### Add filters

Select **Add Filter** to add a filter row. In each row:

1. Choose a column.
2. Choose an operator.
3. Enter a value when the operator requires one.

Use **Remove** on a row to delete that filter, or **Clear** to remove all filters. All filter rows are combined with `AND`. You may filter on a column that is not selected for output, and you may add more than one filter for the same column.

Available operators depend on the column type:

| Column type | Available operators |
| --- | --- |
| Text | equals, does not equal, contains, starts with; is null and is not null when allowed |
| Number or date/time | `=`, `<>`, `>`, `>=`, `<`, `<=`; is null and is not null when allowed |
| True/False or GUID | equals, does not equal; is null and is not null when allowed |

Some specialist column types can be displayed but cannot be filtered. For **contains** and **starts with**, `%`, `_`, and `[` are treated as ordinary characters rather than wildcard syntax.

### Preview, export, or continue in Query

- **Preview** retrieves up to the **Preview row limit** from Settings. Preview results are unordered, so the rows may differ between runs. A limit reduces the number of returned rows but does not guarantee that a large or unindexed table will respond quickly.
- **Export Preview** saves exactly the rows and columns currently displayed. It does not run the preview again and does not export the complete table.
- **Send to Query** creates editable SQL from the current table, checked columns, and filters, then switches to the Query tab. It does not execute the SQL. If the Query editor already contains text, you must confirm before it is replaced.

The displayed preview and the current selections are independent. Changing the table, output columns, filters, or preview limit does not change the displayed preview until **Preview** succeeds again. **Export Preview** uses the displayed preview; **Send to Query** uses the current selections.

## Settings tab

Settings apply to all connections. Change a value and select **Save Settings**; the new values become active only after the save succeeds.

| Setting | Allowed range | Default | What it controls |
| --- | ---: | ---: | --- |
| **Preview row limit** | 10–500 | 100 | Maximum rows returned by each Data Explorer preview |
| **Maximum unordered rows** | 100–2,000 | 1,000 | Maximum rows retained for a Query result without top-level `ORDER BY` |
| **Result data limit (MiB)** | 128–1,024 | 256 | Allocation budget for retained Query pages/results and Data Explorer previews |
| **Query/Export timeout (seconds)** | 5–3,600 | 120 | Time allowed for previews, queries, counts, paging requests, and Excel export |

The 10-second connection timeout is fixed and is not changed by the Query/Export timeout.

## Important usage notes

- SQL Utility is read-only, but queries can still be expensive. Filters, joins, sorting, grouping, and counting may scan or process large amounts of data before the row limit is applied.
- Work runs synchronously. The window may temporarily appear unresponsive while SQL Server or Excel export work is in progress.
- SQL text, results, Data Explorer selections, filters, previews, and export destinations are not saved when the application closes.
- Only one active connection, Query result, and Data Explorer preview are kept at a time.
- If a retained Query page/result or Data Explorer preview reaches the configured result data limit, the operation stops without displaying a partial result. Review the selected columns and filters before retrying. Actual process memory can be higher than this limit because the application and Windows controls have their own overhead.
- Query export produces one Excel worksheet. A single text value longer than Excel's 32,767-character cell limit or a result larger than 1,048,575 data rows stops the export with an error instead of creating a partial workbook.

## Troubleshooting

- **The application does not start:** confirm that all eight runtime files are present in the required structure and start it with `StartSqlUtility.cmd` on Windows PowerShell 5.1. If an integrity-check message appears, extract a fresh copy of the original package instead of editing the distributed files.
- **The connection fails:** verify the server or alias, database name, network/Citrix access, and your Windows-account permissions. The application does not support SQL usernames and passwords.
- **Saved connections or settings disappear:** move the complete application to a folder where you have write access, then test the connection or save the settings again.
- **A preview, query, count, or export times out:** narrow the data, use indexed filters, add a stable `ORDER BY` where appropriate, or increase the Query/Export timeout in Settings.
- **A result reaches the data limit:** select fewer columns, exclude large text or binary columns, or add filters. Increase the Result data limit only when the Citrix session has sufficient memory.
- **Export is disabled:** execute the current editor text again if it changed, or add `ORDER BY` when an unordered result was truncated.
