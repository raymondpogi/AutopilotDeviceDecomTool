# AutopilotDeviceDecomTool
Device Decom Tool to Automate removal in Intune,Entra,AD,SCCM

# Device Decommission Tool

A PowerShell GUI utility designed for **System Engineers** and **IT Admins** to safely and consistently decommission or reissue Windows devices.  
It integrates with **Intune (Autopilot)**, **Entra ID (Azure AD)**, **Active Directory**, and optionally **SCCM**.

---

## ✨ Features
- Secure **MFA-based authentication** with delegated rights  
- **Module checker** – validates required PowerShell modules before use  
- **Serial number & device name checks** – prevents accidental deletions  
- Conditional **button activation** (only available if checks succeed)  
- **Triple confirmation prompts** before executing destructive actions  
- **HWID Autopilot upload** (with serial existence protection)  
- **Comprehensive logging** (Date+Time+Action+Device) saved to script location  
- **SCCM removal optional** – automatically disabled if SCCM console not detected  

---

## 🔧 Prerequisites
- Windows 10/11 with **PowerShell 5.1+** or **PowerShell 7+**  
- Service account with delegated rights:
  - **Intune / Autopilot**: `DeviceManagementManagedDevices.ReadWrite.All`, `DeviceManagementServiceConfig.ReadWrite.All`  
  - **Entra ID**: `Directory.ReadWrite.All`  
  - **Active Directory**: Rights to disable & move computer objects  
  - **SCCM (optional)**: RBAC role with device deletion rights  
- Required PowerShell modules:
  - `Microsoft.Graph.Intune`  
  - `Microsoft.Graph.DeviceManagement`  
  - `ActiveDirectory`  
  - `ConfigurationManager` (for SCCM, optional)  

---

## 🚀 Usage
1. **Login** with your service account → MFA will be triggered  
2. Run **Check Modules** → install missing modules if prompted  
3. Run **Check Serial** or **Check Device Name**  
4. Buttons will enable/disable depending on results:  
   - Serial only → Enrollment removal enabled  
   - Device name only → Entra, AD, SCCM removal enabled  
   - Both → All enabled  
5. Click a button → **triple confirmation prompts** appear  
6. Action executes → status shows Success/Fail  
7. Logs saved in script folder automatically  

---

## 📂 Logging
Logs are stored in the script directory with the format:  


---

## ⚠️ Limitations
- Requires **internet access** to Microsoft Graph API  
- SCCM actions only available if SCCM console + module installed  
- Service account password rotation must be managed externally  
- Only tested on **Windows 10/11 Enterprise** with PowerShell 5.1+  

---

## 📖 Documentation
A full **line-by-line explanation guide** is provided in the repository (`.docx`) for both engineers and helpdesk staff.  
It covers:
- Graph API calls and required permissions  
- Function-by-function explanations  
- Triple confirmation workflow  
- Error codes & troubleshooting  

---

## 🤝 Contributing
Pull requests are welcome.  
For major changes, please open an issue first to discuss what you would like to change.  

---

## 📜 License
[MIT](LICENSE)  
