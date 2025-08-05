# ![EventLogHelper Icon](https://raw.githubusercontent.com/dparvin/EventLogHelper/main/Code/EventLogHelper/NuGet.Pack/Images/EventLogHelper-icon-64x64.png) EventLogHelper

**EventLogHelper** is a lightweight .NET library designed to make logging to the Windows Event Log simple, safe, and flexible — without requiring repetitive boilerplate or administrative headaches.

---

## 🚀 Features

- ✅ One-line logging to the Windows Event Log
- ✅ Automatically creates event sources (with safe fallbacks)
- ✅ Works in non-elevated environments
- ✅ Configurable defaults via `App.config` or `Web.config`
- ✅ Gracefully handles permission issues and registry conflicts
- ✅ Supports structured messages and custom log levels

---

## 💡 Why Use EventLogHelper?

The built-in `.NET` `System.Diagnostics.EventLog` API requires:
- Admin rights to create sources
- Manually checking for source/log conflicts
- Verbose, repeated code everywhere you want to log

**EventLogHelper** abstracts all that into a single method call, with smart defaults and self-healing behavior.

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
    writeInitEntry: True)
```

You can also configure default values and plug in a custom writer for unit testing or specialized behavior.

---

## 🔒 Security Notes

- Creating new sources requires admin privileges.
- In non-elevated environments, EventLogHelper will automatically fall back to an existing source (e.g., the log name or "Application").

---

## 📄 License

MIT License. See [LICENSE](LICENSE) for details.
