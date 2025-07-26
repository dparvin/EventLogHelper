# ![PropertyGridHelpers Icon](https://raw.githubusercontent.com/dparvin/EventLogHelper/main/Code/EventLogHelper/NuGet.Pack/Images/EventLogHelper-icon-64x64.png) EventLogHelper

**EventLogHelper** is a lightweight .NET library designed to make logging to the Windows Event Log simple, safe, and flexible — without requiring repetitive boilerplate or administrative headaches.

## Status
[![Coverage](https://raw.githubusercontent.com/dparvin/EventLogHelper.CodeCoverage/main/Badges/badge_combined.svg)](https://dparvin.github.io/EventLogHelper/)

[![Build status](https://addtracker.visualstudio.com/Projects/_apis/build/status/EventLogHelper%20Build)](https://addtracker.visualstudio.com/Projects/_build/latest?definitionId=9)

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

## 🛠️ Installation

You can add EventLogHelper to your project by cloning or adding the source file directly (NuGet package coming soon).

---

## 🧾 Example Usage

```csharp
using EventLogHelper;

SmartEventLogger.Log("App started.");
SmartEventLogger.Log("Invalid user input", EventLogEntryType.Warning);
SmartEventLogger.Log("Fatal error", EventLogEntryType.Error);
```

You can optionally specify the source and log:

```csharp
SmartEventLogger.Log("Service initialized.", EventLogEntryType.Information, source: "MyService", logName: "MyCompanyLog");
```

---

## ⚙️ Configuration (App.config)

```xml
<configuration>
  <appSettings>
    <add key="EventLog.DefaultLogName" value="Application" />
    <add key="EventLog.DefaultSource" value="MyApp" />
    <add key="EventLog.FallbackToLogNameOnError" value="true" />
    <add key="EventLog.LogLevel" value="Information" />
  </appSettings>
</configuration>
```

---

## 🔒 Security Notes

- Creating new sources requires admin privileges.
- In non-elevated environments, EventLogHelper will automatically fall back to an existing source (e.g., the log name or "Application").

---

## 📦 Roadmap

- Structured logging support
- Optional fallback to file or ETW

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
