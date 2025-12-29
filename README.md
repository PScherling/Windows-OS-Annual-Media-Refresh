# Windows Offline Media Refresh (LTSC / Server)

Enterprise-grade PowerShell script for **annual offline servicing of Windows installation media** using DISM.

This project automates the refresh of:

- **Windows 10/11 Enterprise LTSC**
- **Windows Server 2021/2025**

by injecting the latest cumulative updates into the base installation media in a **safe, repeatable, auditable, and logged** way.

<img width="1579" height="851" alt="image" src="https://github.com/user-attachments/assets/23b1b2a9-07b5-4390-8b38-9d04f01ddfda" />

My solution is inspired by this article from microsoft: https://learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update

---

## ✨ Features

- **Annual, year-based media isolation**
  - Each refresh is stored under its own year (`base/2025`, `base/2026`, …)
  - Previous years are never overwritten

- **Offline servicing only**
  - Uses DISM against WIM images
  - No online updates, no WSUS dependency

- **Latest LCU only**
  - Injects the most recent cumulative update
  - No chaining of historical updates

- **WinRE-safe servicing**
  - Uses **SafeOS Dynamic Update** for Windows Recovery Environment
  - Avoids unsupported LCU injection into WinRE

- **Enterprise logging**
  - Timestamped logfile per execution
  - Console + file logging with severity levels (INFO / OK / WARN / ERROR)
  - Total runtime measurement

- **Failure-resilient**
  - Automatic cleanup of stale DISM mounts
  - Safe re-runs after failure
  - No reboot required under normal conditions

- **Overwrite protection**
  - Completion marker prevents accidental re-processing
  - Explicit ISO hash validation prevents mixing base media

- **Deployment-ready output**
  - Resulting media can be imported directly into **WDS** or similar solutions

---

## 📁 Repository Structure

```text
mediaRefresh/
├─ iso/
│  └─ Windows_11_LTSC.iso
│
├─ packages/
│  ├─ CU/
│  │  └─ windows11-kbxxxxxxx-x64.msu
│  ├─ SafeOS/
│  │  └─ SafeOS_DU.cab
│  ├─ DotNet/
│  │  └─ dotnet-kbxxxxxxx.msu
│  └─ SSU/
│     └─ ssu-kbxxxxxxx.msu or .cab
│
├─ base/
│  ├─ 2025/
|  |  ├─ Windows_11_LTSC.iso/
│  │  |  ├─ oldMedia/
│  │  |  ├─ newMedia/
│  │  │  ├─ RefreshInfo.json
│  │  │  └─ .refresh_completed
│
├─ log/
│  │  └─ mediaRefresh_2025-12-23_18-42-10.log
│
└─ temp/
```

---

## 🚀 Workflow Overview

1. Create the **main directory** `D:\mediaReFresh\` including subdirectoryíes `iso\`, `packages\`, `packages\CU` etc. like in the structure above 
2. Place the **base ISO** into `iso\`
3. Place the **latest LCU** into `packages\CU\`
4. > In case you need to install dependencies (like for KB5071547, you need KB5030216 to install too), place both .msu files in the directory and update the filename this for example:
   > `packages\CU\1_KB5030216.msu`
   > `packages\CU\2_KB5071547.msu`
   > This guarantees, that the updates will be integrated in the image in the correct order
5. (Optional) Place SafeOS and .NET updates into their folders
6. Run the script **as Administrator**
7. The script will:
   - Mount the ISO
   - Extract it into a year-specific `oldMedia`
   - Create a serviced `newMedia`
   - Update WinRE safely
   - Validate results
   - Log all actions and runtime
   - Lock the year with a completion marker

---

## 🛡️ Safety & Design Guarantees

| Scenario | Behavior |
|--------|----------|
| First run of a year | ✔ Allowed |
| Failure mid-run | ✔ Safe to re-run |
| Completed year | ❌ Overwrite blocked |
| Old year media | ✔ Preserved |
| ISO mismatch | ❌ Hard fail |
| Stale DISM mounts | ✔ Auto-cleaned |

<img width="1366" height="454" alt="image" src="https://github.com/user-attachments/assets/b1901fbc-3d2a-4b87-8fc6-e7796227cab6" />

---

## 📄 Output Artifacts

Each successful run produces:

- Refreshed `install.wim`
- Updated WinRE
- `RefreshInfo.json` (audit metadata)
- `.refresh_completed` marker (overwrite protection)
- Timestamped execution logfile

Example `RefreshInfo.json`:

```json
{
  "Year": 2025,
  "ISO": "Windows_11_LTSC.iso",
  "ISO_SHA256": "ABCDEF123456...",
  "LCU": "windows11-kbxxxxxxx-x64.msu",
  "Completed": "2025-12-23T18:42:10",
  "Runtime": "02:46:15"
}
```

---

## 🔐 Requirements

- Windows Server or Windows Client with DISM
- PowerShell 5.1+
- Administrative privileges
- Local NTFS storage (ISO mounting required)

---

## ⚠️ Notes

- The script intentionally **refuses to overwrite completed yearly media**
- To re-run a completed year, the marker file must be removed manually
- SafeOS updates are optional but recommended for WinRE consistency

---

## 👤 Author

**Author:** Patrick Scherling  
**Contact:** @Patrick Scherling  

---

> ⚡ *“Automate. Standardize. Simplify.”*  
> Part of Patrick Scherling’s IT automation suite for modern Windows Server infrastructure management.
