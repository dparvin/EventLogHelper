# ![EventLogHelper Icon](https://raw.githubusercontent.com/dparvin/EventLogHelper/main/Code/EventLogHelper/NuGet.Pack/Images/EventLogHelper-icon-64x64.png) EventLogHelper

**EventLogHelper** is a lightweight .NET library designed to make logging to the Windows Event Log 
simple, safe, and flexible — without requiring repetitive boilerplate or administrative headaches.

## Status
[![Coverage](https://raw.githubusercontent.com/dparvin/EventLogHelper.CodeCoverage/main/Badges/badge_combined.svg)](https://dparvin.github.io/EventLogHelper/)

[![Build status](https://addtracker.visualstudio.com/Projects/_apis/build/status/EventLogHelper%20Build)](https://addtracker.visualstudio.com/Projects/_build/latest?definitionId=9)

---

## 🚀 Features

- ✅ One-line logging to the Windows Event Log
- ✅ Automatically creates event sources (with safe fallbacks)
- ✅ Works in non-elevated environments
- ✅ Configurable defaults via App.config, Web.config, or appsettings.json (automatically loaded at first use)
- ✅ Gracefully handles permission issues and registry conflicts
- ✅ Configurable logging severity (filter logs by importance)

---

## 💡 Why Use EventLogHelper?

The built-in `.NET` `System.Diagnostics.EventLog` API requires:
- Admin rights to create sources
- Manually checking for source/log conflicts
- Verbose, repeated code everywhere you want to log

**EventLogHelper** abstracts all that into a single method call, with smart defaults and self-healing 
behavior.

---

## 🛠️ Installation

Install the NuGet package:

```bash

dotnet add package EventLogHelper

```

Or via the NuGet Package Manager Console:

```powershell

Install-Package EventLogHelper

```

See the NuGet Gallery page for more details.

---

## 🧾 Example Usage

Basic logging with default source and log name:

```vbnet

SmartEventLogger.Log("Application started.")
SmartEventLogger.Log("A warning occurred.", EventLogEntryType.Warning)
SmartEventLogger.Log("An error occurred.", EventLogEntryType.Error)

```

Custom source and log name:

```vbnet

SmartEventLogger.MachineName = "."
SmartEventLogger.LogName = "MyCompanyLog"
SmartEventLogger.SourceName = "MyServiceSource"

SmartEventLogger.Log("Service initialized.",
                     EventLogEntryType.Information)

```

Or with a fluent interface using `GetLog`:
```vbnet

SmartEventLogger.GetLog("MyCompanyLog", "MyServiceSource").
    LogEntry("Service initialized.", EventLogEntryType.Information)

```

Advanced usage with full customization:

```vbnet

SmartEventLogger.Log(
    _machineName: ".",
    _logName: "CustomLog",
    sourceName: "CustomSource",
    message: "This is a custom log entry.",
    eventType: EventLogEntryType.Information,
    eventID: 1001,
    category: 0,
    rawData: Nothing,
    maxKilobytes: 1024 * 1024,    ' 1 MB
    retentionDays: 7,
    writeInitEntry: True, 
    EntrySeverity: LoggingSeverity.Info)

```

## 🔑 How the Levels Work

LoggingLevel (global threshold) – sets how strict the logger is.

None → no logs are written

Verbose → everything is written

Critical → only critical logs are written

LoggingSeverity (per-entry importance) – sets how important a log entry is considered.

Verbose (least important) through Critical (most important)

A log entry is written only if:

```vbnet

    EntrySeverity >= CurrentLoggingLevel

```

---

```vbnet

' Configure defaults
SmartEventLogger.CurrentLoggingLevel = LoggingLevel.Normal   ' write Info, Warning, Error, Critical
SmartEventLogger.LoggingSeverity = LoggingSeverity.Info      ' default for entries without explicit severity

' Explicit severity: written (Warning >= Normal)
SmartEventLogger.Log("Low disk space.",
                     EventLogEntryType.Warning,
                     EntrySeverity:=LoggingSeverity.Warning)

' Explicit severity: skipped (Verbose < Normal)
SmartEventLogger.Log("Debug trace here...",
                     EventLogEntryType.Information,
                     EntrySeverity:=LoggingSeverity.Verbose)

' Implicit severity: uses default (Info >= Normal, so written)
SmartEventLogger.Log("Application started.",
                     EventLogEntryType.Information)

```

You can also configure default values and plug in a custom writer for unit testing or specialized 
behavior.

---
## 🔄 Source Resolution Behavior

Sometimes an event source exists, but it’s registered under a different log than the one you want to 
write to. By default, this mismatch causes an exception. You can control how EventLogHelper resolves 
these situations with the SourceResolutionBehavior property:

```vbnet
Enum SourceResolutionBehavior
    Strict            ' Fail if the source is tied to a different log
    UseSourceLog      ' Automatically switch to the log the source is registered under
    UseLogsDefaultSource ' Use the default source for the requested log instead
End Enum
```

### Example

```vbnet
' Configure global behavior
SmartEventLogger.SourceResolutionBehavior = SourceResolutionBehavior.UseSourceLog

' Case 1: Source is under the wrong log
'   Strict → throws InvalidOperationException
'   UseSourceLog → log is switched automatically
'   UseLogsDefaultSource → original source is preserved in the message text,
'                          but the actual log entry uses the log’s default source
```

This allows you to balance safety (strict checking) against robustness (automatic fallback), depending 
on your deployment environment.

---

## ⚙️ Configuration via App Settings

SmartEventLogger can initialize itself automatically using application configuration files:

- .NET Framework – settings are read from <appSettings> in App.config or Web.config.
- .NET Core / .NET 5+ – settings are read from appsettings.json.

Recognized keys:

| Key                                       | Description                                               |
| ----------------------------------------  | --------------------------------------------------------- |
| `EventLogHelper.MachineName`              | Target machine name for logging (default: local machine)  |
| `EventLogHelper.LogName`                  | Event log name (default: Application)                     |
| `EventLogHelper.SourceName`               | Event source name                                         |
| `EventLogHelper.MaxKilobytes`             | Maximum log size in KB                                    |
| `EventLogHelper.RetentionDays`            | Days to retain log entries                                |
| `EventLogHelper.WriteInitEntry`           | Whether to write an initialization entry (`true`/`false`) |
| `EventLogHelper.TruncationMarker`         | String marker for truncated messages                      |
| `EventLogHelper.ContinuationMarker`       | String marker for continued messages                      |
| `EventLogHelper.AllowMultiEntryMessages`  | Whether to split long messages into multiple entries      |
| `EventLogHelper.LoggingSeverity`          | The default logging severity level for logging where the severity is not specified. |
| `EventLogHelper.CurrentLoggingLevel`      | Current logging level (defaults to `Normal`). This can be set to control the verbosity of logs written. Log Entries with a severity below this level are not written. |
| `EventLogHelper.SourceResolutionBehavior` | How to handle mismatched source and log names (defaults to `Strict`). |
| `EventLogHelper.IncludeSourceInMessage`   | Whether to include the source name in the log message. (defaults to `true`) |

Example – App.config

```xml
<appSettings>
    <add key="EventLogHelper.LogName" value="Application" />
    <add key="EventLogHelper.SourceName" value="MyAppSource" />
</appSettings>
```

Example – appsettings.json

```json
{
  "EventLogHelper.LogName": "Application",
  "EventLogHelper.SourceName": "MyAppSource"
}
```

If you don’t call InitializeConfiguration() explicitly, these settings will be loaded automatically 
the first time you use the logger.

---

## ⚡ EventLogHelper.Elevator

Some operations in Windows Event Log require administrator rights, specifically 
when creating new event logs or registering new sources.

Normal logging (writing entries to an existing log/source) works without elevation.

To handle this securely and transparently, EventLogHelper ships with a small helper 
application:

```EventLogHelper.Elevator.exe```

- Runs only when needed (creating a new log or source).
- Automatically triggers a UAC prompt asking for elevated permissions.
- Configures the log size and retention policy only at creation time.
- If the user declines elevation, EventLogHelper will safely fall back to writing entries under the default Application log and source.

### 🔑 Important Notes

- You don’t run this helper directly — it’s invoked automatically by EventLogHelper.
- If you open Event Viewer while creating a new log, you may need to close and reopen it to see the new log appear.
- Log names must not contain reserved characters (\ / * ?) or spaces at the beginning or end.
- Custom sources can be added later without reconfiguring the log.

---

## 🔒 Security Notes

- Creating new sources requires admin privileges.
- In non-elevated environments, EventLogHelper will automatically fall back to an existing source 
- (e.g., the log name or "Application").

---

## 📦 Roadmap

- Optional fallback to file or ETW

---

## 🙌 Contributions

Contributions are welcome! Please file issues or open pull requests to suggest improvements, report 
bugs, or extend functionality.

---

## 📄 License

MIT License. See [LICENSE](LICENSE) for details.

---

## ✨ Credits

Created and maintained by [@dparvin](https://github.com/dparvin).

---

### Related Locations
* [Code Coverage Report](https://dparvin.github.io/EventLogHelper/)
* [Published Package](https://www.nuget.org/packages/EventLogHelper)
