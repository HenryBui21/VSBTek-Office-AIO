# Release Notes

## Version 1.0 - Initial Release 🎉
**Release Date:** October 25, 2025  
**Author:** VSBTek (Henry Bui)

---

### 🎯 Overview

VSBTek Office AIO (All-In-One) Toolkit - Công cụ cài đặt và quản lý Microsoft Office tự động với giao diện dễ sử dụng.

---

### ✨ Features

#### **1. Office Installation (Menu 1)**
- ✅ Support 8 versions: Office 2007, 2010, 2013, 2016, 2019, 2021, **2024**, 365
- ✅ Choose 32-bit or 64-bit architecture
- ✅ Select specific apps to install
- ✅ Quick install mode (one-click for basic apps)
- ✅ Custom install mode (choose individual apps)

#### **2. Project & Visio Installation (Menu 2)**
- ✅ Support Project Pro: 2016, 2019, 2021
- ✅ Support Visio Pro: 2016, 2019, 2021
- ✅ Retail/Volume options

#### **3. Export Office Shortcuts (Menu 3)**
- ✅ Auto-search shortcuts from multiple locations
- ✅ Copy all Office apps to Desktop
- ✅ Support both old (2007-2013) and new (2016+) Office

#### **4. Uninstall Office Completely (Menu 4)**
- ✅ **OfficeScrubber** (Recommended) - Official Microsoft tool
  - Remove all Office versions (2003-2024 + 365)
  - Clean License, Registry, Files
  - Support both MSI and Click-to-Run
- ✅ **Revo Uninstaller** (Backup) - Alternative uninstaller

#### **5. Download Office ISO (Menu 5)**
- ✅ Office 2019 Pro Plus
- ✅ Office 2021 Pro Plus
- ✅ Office 2024 Pro Plus ⭐ NEW
- ✅ Office 365 Pro Plus
- ✅ Direct links from **Microsoft CDN** (safe & fast)

---

### 🆕 What's New in v1.0

#### **Office 2024 Support** ⭐
- Added Office 2024 Pro Plus installation
- Auto-skip Publisher (removed by Microsoft in Office 2024)
- Full support for Word, Excel, PowerPoint, Outlook, OneNote, Access, OneDrive

#### **Improvements**
- ✅ Updated Microsoft CDN download links
- ✅ Improved UI/UX with clear menus
- ✅ Fixed Vietnamese encoding issues
- ✅ Better error handling
- ✅ Optimized script performance
- ✅ Removed duplicate tools

#### **Bug Fixes**
- 🐛 Fixed Office 365 installation issues
- 🐛 Fixed Product ID detection for Office 2021/2024
- 🐛 Fixed Publisher detection for Office 2024
- 🐛 Improved Configuration.xml generation

---

### 📋 System Requirements

- **OS:** Windows 10/11 (recommended)
  - Windows 7/8/8.1: Limited support (Office 2010, 2013, 2016 Volume only)
- **Architecture:** 32-bit and 64-bit
- **Administrator Rights:** Required
- **Internet Connection:** Required for downloading Office

---

### 📦 What's Included

```
VSBTek-Office-AIO/
├── C2R.bat                 # Main launcher
├── README.md               # Documentation
├── LICENSE                 # MIT License
└── xml/
    ├── menu.cmd            # Main menu
    ├── office.cmd          # Office installer
    ├── project_visio.cmd   # Project/Visio installer
    ├── setup.exe           # Click-to-Run installer
    └── remove_office_tan_goc/
        └── OfficeScrubber/ # Microsoft official uninstaller
```

---

### 🚀 Quick Start

1. **Download** the release package
2. **Extract** to any folder
3. **Run** `C2R.bat` as Administrator
4. **Choose** from menu options (1-5)
5. **Follow** on-screen instructions

---

### 📝 Office Versions Comparison

| Feature | 2007 | 2010 | 2013 | 2016 | 2019 | 2021 | 2024 | 365 |
|---------|------|------|------|------|------|------|------|-----|
| 32/64-bit | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Core Apps | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Publisher | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ | ✅ |
| OneDrive | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ | ✅ | ✅ |
| Teams | ❌ | ❌ | ❌ | ❌ | ❌ | ❌ | ❌ | ✅ |
| Project/Visio | ❌ | ✅ | ✅ | ✅ | 🔧 | 🔧 | ❌ | 🔧 |

**Legend:** ✅ = Supported | ❌ = Not available | 🔧 = Separate install (Menu 2)

---

### ⚠️ Important Notes

1. **Office 2024:** Publisher is removed by Microsoft (not a bug)
2. **Project/Visio:** For Office 2019/2021/365, install separately via Menu 2
3. **License:** Ensure you have valid Microsoft Office license
4. **Backup:** Backup your data before uninstalling Office

---

### 🔗 Links

- **GitHub Repository:** https://github.com/HenryBui21/VSBTek-Office-AIO
- **Issues:** https://github.com/HenryBui21/VSBTek-Office-AIO/issues
- **Documentation:** [README.md](README.md)

---

### 👤 Credits

- **Original Author:** Thanos
- **Rebuild & Improvements:** VSBTek (Henry Bui)
- **OfficeScrubber:** Microsoft Support and Recovery Assistant (SaRA)
- **Microsoft Office:** Microsoft Corporation

---

### 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

### 🙏 Thank You

Thank you for using VSBTek Office AIO Toolkit!

If you find this tool useful, please give it a ⭐ on GitHub!

---

**VSBTek Office AIO v1.0** | Built with ❤️ by VSBTek

