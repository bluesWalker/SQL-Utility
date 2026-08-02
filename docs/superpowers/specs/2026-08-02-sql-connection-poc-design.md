# SQL Connection GUI POC Design

## Purpose

Build a minimal, portable Windows GUI that proves a Citrix-hosted SQL Server can be reached with the signed-in user's Windows identity. The POC must accept connection details at runtime, execute one fixed read-only diagnostic query, and display its result without persisting user input or runtime data.

## Scope

The POC will:

- Start from a double-clickable command launcher or directly from PowerShell.
- Ask the user to enter the SQL Server and database names into initially blank fields.
- Connect with Windows integrated authentication.
- Execute one built-in read-only diagnostic query.
- Display the returned row in a read-only GUI grid.
- Show validation, connection, and query errors in the GUI.

The POC will not:

- Accept arbitrary SQL.
- Ask for or store a username or password.
- Save server names, database names, results, settings, history, logs, exports, cache files, or temporary files.
- Install software, compile an executable, or modify the registry.

## Runtime and Dependencies

The target runtime is Windows PowerShell 5.1 in the Citrix environment. The implementation uses only Windows and .NET Framework components normally present with that runtime:

- `System.Windows.Forms` for the GUI.
- `System.Drawing` for basic control sizing and layout.
- `System.Data.SqlClient` for SQL Server connectivity.

The application does not depend on `sqlcmd`, third-party PowerShell modules, NuGet packages, Python, Office, administrator rights, or an installer. Operational prerequisites are network and name-resolution access to the SQL Server plus database permission for the current Windows account.

## Deliverables

- `StartSqlPoc.cmd`: a small launcher that starts the PowerShell script from its own directory without changing execution policy permanently.
- `SqlConnectionPoc.ps1`: the complete GUI and database logic.

Both files are portable text files that can be copied together to the cloud drive. If command-file execution is restricted, the PowerShell script can be started from an already permitted PowerShell session.

## User Interface

The window contains:

- An initially blank Server field. It accepts standard SQL Server data-source forms such as `server`, `server\\instance`, or `server,port`.
- An initially blank Database field.
- A **Test connection and query** button.
- A status label for progress and outcomes.
- A read-only result grid.

The interface contains no username, password, query editor, save, export, recent-connection, or settings controls.

## Connection and Query Flow

1. The user enters a server and database.
2. The application rejects blank or whitespace-only values before attempting a connection.
3. It constructs a `SqlConnectionStringBuilder` in memory. The builder sets the entered data source and initial catalog, enables integrated security, identifies the application, and applies a short connection timeout. It never includes SQL credentials.
4. The application clears any previous result and reports that the test is in progress.
5. It opens a `SqlConnection` and executes this fixed query with a command timeout:

   ```sql
   SELECT
       @@SERVERNAME AS ServerName,
       DB_NAME() AS DatabaseName,
       SYSTEM_USER AS LoginName,
       GETDATE() AS ServerTime;
   ```

6. A `SqlDataAdapter` loads the result into an in-memory `DataTable`.
7. The application binds the table to a read-only `DataGridView` and reports success.
8. Connections and commands are disposed deterministically after success or failure.

The query reads server/session metadata only and does not change database state.

## Error Handling

- Missing server or database values produce a concise validation message without a connection attempt.
- SQL, network, name-resolution, authentication, and permission failures are caught and displayed in the window.
- The status distinguishes success from failure and the test button is restored after each attempt.
- No error details are written to disk.
- Closing the window discards all entered values, connection objects, status text, and results.

## Security and Persistence

- Authentication uses only the Windows identity of the process through integrated security.
- Connection-string construction uses `SqlConnectionStringBuilder`, so field contents cannot inject additional connection-string properties.
- The fixed query cannot be modified through the GUI.
- The application deliberately performs no file, registry, environment-variable, clipboard, or network writes other than the SQL Server connection and query protocol.
- Server and database names are absent from both source files and start blank on every launch.

The underlying Windows, PowerShell, .NET, Citrix, or SQL Server environment may maintain its own administrative or security auditing; the POC itself creates no persistence.

## Verification and Acceptance

Local verification will confirm:

- The PowerShell script parses under Windows PowerShell.
- Required .NET assemblies and types load.
- Both input fields are blank by default.
- The connection uses integrated security and contains no credential fields.
- The SQL text is the approved fixed diagnostic query.
- The source contains no persistence or export operations.
- The launcher resolves the PowerShell script relative to its own location.

Final acceptance must be performed inside Citrix because the local development machine cannot resolve the host alias. The POC succeeds when a user can copy both files to the cloud drive, launch the GUI, enter the server and database, click the test button, and see the four diagnostic columns populated for the connected SQL Server session.

## Explicitly Deferred Work

Arbitrary query editing, multiple result sets, cancellation, asynchronous execution, result export, saved connections, history, syntax highlighting, transaction controls, and packaging as an executable are deferred until this connectivity POC succeeds in Citrix.
