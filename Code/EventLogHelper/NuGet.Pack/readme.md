# ![EventLogHelper Icon](https://raw.githubusercontent.com/dparvin/EventLogHelper/main/Code/EventLogHelper/NuGet.Pack/Images/EventLogHelper-icon-64x64.png) EventLogHelper

**EventLogHelper** is a lightweight .NET library designed to make logging to the Windows Event Log 
simple, safe, and flexible — without requiring repetitive boilerplate or administrative headaches.

---

## 🚀 Features

- ✅ One-line logging to the Windows Event Log
- ✅ Automatically creates event sources (with safe fallbacks)
- ✅ Works in non-elevated environments
- ✅ Configurable defaults via App.config, Web.config, or appsettings.json (automatically loaded at 
first use)
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
    maxKilobytes: 1024 * 1024,    ' 1 GB
    retentionDays: 7,
    writeInitEntry: True,
    EntrySeverity: LoggingSeverity.Info)

```

You can also configure default values and plug in a custom writer for unit testing or specialized 
behavior.

---

## ⚙️ Configuration via App Settings

SmartEventLogger can initialize itself automatically using application configuration files:

- .NET Framework – settings are read from <appSettings> in App.config or Web.config.
- .NET Core / .NET 5+ – settings are read from appsettings.json.

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
(e.g., the log name or "Application").

---

## 📄 License

MIT License. See [LICENSE](LICENSE) for details.
