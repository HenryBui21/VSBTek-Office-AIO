# Release Notes - Version 1.1

## Version 1.1 - Bug Fixes & Improvements 🔧
**Release Date:** October 25, 2025  
**Author:** VSBTek (Henry Bui)

---

## 🎯 What's New in v1.1

### 🐛 Critical Bug Fixes

#### **1. Fixed Office 2016 Click-to-Run Support** ⭐ IMPORTANT
- **Issue:** Office 2016 was incorrectly forced to download mode
- **Fix:** Office 2016 now supports Click-to-Run installation (like 2019/2021/2024/365)
- **Impact:** Faster installation, no need to download large ISO files
- **Benefit:** Users can now install Office 2016 directly from internet with app selection

#### **2. Fixed App Summary Display**
- **Issue:** Selected apps display was split across multiple lines
- **Fix:** All selected apps now display on a single line
- **Before:**
  ```
  Bao gom: Word + Excel + PowerPoint + Access + Publisher + Outlook
          + OneNote + OneDrive + Teams
  ```
- **After:**
  ```
  Bao gom: Word + Excel + PowerPoint + Access + Publisher + Outlook + OneNote + OneDrive + Teams
  ```

#### **3. Fixed Variable Mapping for Apps**
- **Issue:** OneDrive and Teams were using wrong variables
- **Fix:** Corrected variable assignments:
  - `h` = OneDrive (was incorrectly set to Teams)
  - `i` = Teams (was using wrong variable `k`)
- **Impact:** Summary now shows correct apps selected by user

#### **4. Updated Quick Install Mode**
- **Changed:** Quick install now includes Teams instead of OneDrive
- **New apps:** Word + Excel + PowerPoint + Outlook + Teams
- **Reason:** Better for collaboration-focused users
- **Note:** OneDrive can still be selected in custom mode

---

## 📋 Detailed Changes

### Code Improvements

**office.cmd:**
- Line 83: Office 2016 now uses `goto:quickselect` instead of `goto:1`
- Line 135: Removed incorrect download redirect for Office 2016
- Line 252: Fixed OneDrive variable from `k` to `h`
- Line 265: Fixed Teams variable from `k` to `i`
- Line 362: Combined app display into single line
- Line 151-157: Updated quick install apps list

**README.md:**
- Updated quick install description
- Fixed typo: "Go cai" → "Go bo"

**RELEASE_NOTES.md:**
- Updated quick install feature description

---

## 🆚 Comparison: v1.0 vs v1.1

| Feature | v1.0 | v1.1 |
|---------|------|------|
| Office 2016 C2R | ❌ Download only | ✅ Full C2R support |
| App summary display | ❌ Multiple lines | ✅ Single line |
| Variable mapping | ❌ Incorrect | ✅ Correct |
| Quick install apps | OneDrive | Teams |

---

## 🎯 Click-to-Run Support Status

| Office Version | C2R Support |
|:---------------|:------------|
| Office 2007 | ❌ Download only (MSI) |
| Office 2010 | ❌ Download only (MSI) |
| Office 2013 | ❌ Download only (MSI) |
| **Office 2016** | ✅ **NOW SUPPORTED** ⭐ |
| Office 2019 | ✅ Supported |
| Office 2021 | ✅ Supported |
| Office 2024 | ✅ Supported |
| Office 365 | ✅ Supported |

---

## 🚀 Upgrade from v1.0

### What You Need to Know

1. **Office 2016 users:** You can now install directly without downloading ISO
2. **Better UX:** App summary is cleaner and easier to read
3. **More accurate:** Selected apps display correctly
4. **Quick install:** Now includes Teams for better collaboration

### Migration Guide

If you're upgrading from v1.0:
1. Download the new release
2. No configuration changes needed
3. All existing features work the same
4. Office 2016 users will see new installation options

---

## 📦 Installation

Same as v1.0:
1. **Download** the release package
2. **Extract** to any folder
3. **Run** `C2R.bat` as Administrator
4. **Choose** from menu options (1-5)

---

## 🔗 Links

- **GitHub Repository:** https://github.com/HenryBui21/VSBTek-Office-AIO
- **Issues:** https://github.com/HenryBui21/VSBTek-Office-AIO/issues
- **Previous Release (v1.0):** https://github.com/HenryBui21/VSBTek-Office-AIO/releases/tag/v1.0

---

## 📝 Full Changelog

```
v1.1 (2025-10-25)
- [FIX] Enable Click-to-Run installation for Office 2016
- [FIX] Display all apps on single line in summary
- [FIX] Correct OneDrive variable assignment
- [FIX] Correct Teams variable assignment
- [UPDATE] Change quick install from OneDrive to Teams
- [FIX] README typo: "Go cai" → "Go bo"
- [UPDATE] Documentation updates
```

---

## ⚠️ Known Issues

None reported.

---

## 🙏 Thank You

Thank you for using VSBTek Office AIO Toolkit and for reporting issues!

Special thanks to users who tested and provided feedback.

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

**VSBTek Office AIO v1.1** | Built with ❤️ by VSBTek

**Download:** [Release v1.1](https://github.com/HenryBui21/VSBTek-Office-AIO/releases/tag/v1.1)

