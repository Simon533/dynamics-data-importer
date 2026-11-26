# Dynamics Data Importer

## Configuration

Set these environment variables before running the application:

- `DATAVERSE_URL`: The URL of the Dataverse organization to connect to.
- `PROFESSION_EXCEL_FILE`: The path to the Excel workbook containing profession records.

For example, in PowerShell:

```powershell
$env:DATAVERSE_URL = "https://your-org.crm4.dynamics.com"
$env:PROFESSION_EXCEL_FILE = "C:\path\to\professions.xlsx"
dotnet run
```

The application uses OAuth authentication through the Dataverse client.
