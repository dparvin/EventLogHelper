# ![EventLogHelper Icon](https://raw.githubusercontent.com/dparvin/EventLogHelper/main/Code/EventLogHelper/NuGet.Pack/Images/EventLogHelper-icon-64x64.png) EventLogHelper

**EventLogHelper** is a lightweight .NET library designed to make logging to the Windows Event Log simple, safe, and flexible — without requiring repetitive boilerplate or administrative headaches.

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
- ✅ Custom log levels

---

## 💡 Why Use EventLogHelper?

The built-in `.NET` `System.Diagnostics.EventLog` API requires:
- Admin rights to create sources
- Manually checking for source/log conflicts
- Verbose, repeated code everywhere you want to log

**EventLogHelper** abstracts all that into a single method call, with smart defaults and self-healing behavior.

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
    writeInitEntry: True)
```

You can also configure default values and plug in a custom writer for unit testing or specialized behavior.

---

## ⚙️ Configuration via App Settings

SmartEventLogger can initialize itself automatically using application configuration files:

.NET Framework – settings are read from <appSettings> in App.config or Web.config.

.NET Core / .NET 5+ – settings are read from appsettings.json.

Recognized keys:

| Key                                      | Description                                               |
| ---------------------------------------- | --------------------------------------------------------- |
| `EventLogHelper.MachineName`             | Target machine name for logging (default: local machine)  |
| `EventLogHelper.LogName`                 | Event log name (default: Application)                     |
| `EventLogHelper.SourceName`              | Event source name                                         |
| `EventLogHelper.MaxKilobytes`            | Maximum log size in KB                                    |
| `EventLogHelper.RetentionDays`           | Days to retain log entries                                |
| `EventLogHelper.WriteInitEntry`          | Whether to write an initialization entry (`true`/`false`) |
| `EventLogHelper.TruncationMarker`        | String marker for truncated messages                      |
| `EventLogHelper.ContinuationMarker`      | String marker for continued messages                      |
| `EventLogHelper.AllowMultiEntryMessages` | Whether to split long messages into multiple entries      |

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

If you don’t call InitializeConfiguration() explicitly, these settings will be loaded automatically the first time you use the logger.

---

## 🔒 Security Notes

- Creating new sources requires admin privileges.
- In non-elevated environments, EventLogHelper will automatically fall back to an existing source (e.g., the log name or "Application").

---

## 📦 Roadmap

- Optional fallback to file or ETW
- Log large messages by paging the message
- configure with `App.config` or `Web.config`
- Setup logging levels so that you can skip logging by severity

---

## 🙌 Contributions

Contributions are welcome! Please file issues or open pull requests to suggest improvements, report bugs, or extend functionality.

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
