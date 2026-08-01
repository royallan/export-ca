# ca-export

Export all **Microsoft Entra Conditional Access (CA) policies** to a readable CSV
overview and/or a full-fidelity JSON file, with GUIDs resolved to display names. Useful
for reviewing, documenting, backing up and diffing policies outside the Entra portal.

## What it does

- Connects to Microsoft Graph and retrieves every Conditional Access policy
- Resolves GUID references (users, groups, directory roles, applications, named
  locations, terms of use) to display names, on demand and cached, so it does not
  download the entire directory
- Flags references that no longer resolve: `(deleted)` for objects found in the recycle
  bin, `(not found)` for stale references, `(app not found)` for apps without a service
  principal
- Writes the output in the format you choose: CSV, JSON, or both

## Output

| File | Purpose |
|------|---------|
| `ConditionalAccess-overview_*.csv` | Flattened, human-readable overview, one row per policy. |
| `ConditionalAccess-policies_*.json` | Raw Graph representation of each policy. Suitable for backup, diffing and re-import. |

The CSV is the readable overview. The JSON is the source of truth: it keeps everything
the CSV flattens away, including full session controls, grant control operator, template
IDs and timestamps.

## Requirements

- **PowerShell 7 or later.**
- The **Microsoft.Graph.Authentication** module (this is the only module needed;
  everything runs through `Invoke-MgGraphRequest`):
  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```
- Delegated scopes, consented on first run: `Policy.Read.All`, `Directory.Read.All`,
  `Agreement.Read.All` (the last is only needed to resolve Terms of Use names).

## Usage

```powershell
.\ca-export.ps1                                   # both CSV and JSON (default)
.\ca-export.ps1 -OutputFormat JSON                # only JSON
.\ca-export.ps1 -OutputFormat CSV -CsvDelimiter ";"   # CSV tuned for Norwegian Excel
.\ca-export.ps1 -OutputDirectory C:\exports       # write somewhere specific
```

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-OutputFormat` | `Both` | `CSV`, `JSON`, or `Both`. |
| `-OutputDirectory` | current folder | Where to write the output files. |
| `-CsvDelimiter` | `,` | CSV delimiter. Use `;` where Excel expects a semicolon (for example Norwegian). |
| `-UseDeviceCode` | off | Sign in with device code flow instead of the interactive browser. Useful on a headless host or if interactive sign-in fails with a broker error. |

## Notes

Multi-value fields (users, groups, apps, and so on) are joined with ` | ` so they do not
collide with the CSV delimiter.

`(not found)` in a user or group column means the reference resolves to neither an active
object nor one in the recycle bin. In practice that is a stale reference to something
that has been removed, and is worth cleaning up in the policy.

## Troubleshooting

**`Connect-MgGraph` fails with "Method not found ... WithLogging"**

This is an assembly clash, not a script bug. It happens when a module that loads an older
MSAL assembly (most often ExchangeOnlineManagement via `Connect-ExchangeOnline`) has
already run in the same session. Run this script from a **fresh** PowerShell window, or
connect Microsoft Graph before Exchange Online. Keeping this tool in its own session
avoids the problem entirely.

If interactive sign-in fails with a broker or window-handle error, use `-UseDeviceCode`.

## Disclaimer

Provided as-is, without warranty. You are responsible for complying with applicable law
and your organization's authorization requirements when exporting policy data.

## License

Released under the MIT License. Free to use, modify and distribute.
