# Release Instructions

## 📦 Hướng Dẫn Tạo Release Trên GitHub (Không Dùng Tag)

---

## 🎯 Release v1.0 - Initial Release

### Bước 1: Vào Trang Releases
1. Truy cập: https://github.com/HenryBui21/VSBTek-Office-AIO/releases
2. Click nút **"Draft a new release"**

### Bước 2: Điền Thông Tin Release

**Choose a tag:**
- Click **"Choose a tag"** 
- Gõ: `v1.0`
- Click **"+ Create new tag: v1.0 on publish"**

**Target:**
- Chọn branch: `main`
- Hoặc chọn commit cụ thể: `2506053` (Release v1.0: Add release notes and fix README typo)

**Release title:**
```
VSBTek Office AIO v1.0 - Initial Release 🎉
```

**Description:**
Copy toàn bộ nội dung từ file: [`docs/RELEASE_NOTES_v1.0.md`](RELEASE_NOTES_v1.0.md)

### Bước 3: Publish
- **KHÔNG tích:** "Set as a pre-release"
- **KHÔNG tích:** "Set as the latest release" (vì v1.1 mới hơn)
- Click **"Publish release"**

---

## 🔧 Release v1.1 - Bug Fixes & Improvements

### Bước 1: Vào Trang Releases
1. Truy cập: https://github.com/HenryBui21/VSBTek-Office-AIO/releases
2. Click nút **"Draft a new release"** (lần 2)

### Bước 2: Điền Thông Tin Release

**Choose a tag:**
- Click **"Choose a tag"**
- Gõ: `v1.1`
- Click **"+ Create new tag: v1.1 on publish"**

**Target:**
- Chọn branch: `main`
- Hoặc chọn commit cụ thể: `a086dc6` (Release v1.1: Bug fixes and improvements)

**Release title:**
```
VSBTek Office AIO v1.1 - Bug Fixes & Improvements 🔧
```

**Description:**
Copy toàn bộ nội dung từ file: [`docs/RELEASE_NOTES_v1.1.md`](RELEASE_NOTES_v1.1.md)

### Bước 3: Publish
- **KHÔNG tích:** "Set as a pre-release"
- **✅ TÍCH:** "Set as the latest release" (đây là version mới nhất)
- Click **"Publish release"**

---

## 📋 Commit Hash Reference

| Version | Commit Hash | Commit Message |
|:--------|:------------|:---------------|
| **v1.0** | `2506053` | Release v1.0: Add release notes and fix README typo |
| **v1.1** | `a086dc6` | Release v1.1: Bug fixes and improvements |

---

## ✅ Checklist

### Release v1.0
- [ ] Vào releases page
- [ ] Tạo tag v1.0
- [ ] Chọn commit `2506053`
- [ ] Copy nội dung từ RELEASE_NOTES_v1.0.md
- [ ] KHÔNG set as latest
- [ ] Publish

### Release v1.1
- [ ] Vào releases page
- [ ] Tạo tag v1.1
- [ ] Chọn commit `a086dc6`
- [ ] Copy nội dung từ RELEASE_NOTES_v1.1.md
- [ ] ✅ Set as latest
- [ ] Publish

---

## 📸 Screenshot Mẫu

**Step 1: Choose tag**
```
[Choose a tag ▼]
v1.0
  + Create new tag: v1.0 on publish
```

**Step 2: Target**
```
Target: main @ a086dc6
```

**Step 3: Release title**
```
VSBTek Office AIO v1.1 - Bug Fixes & Improvements 🔧
```

**Step 4: Checkboxes**
```
☐ Set as a pre-release
☑ Set as the latest release  (only for v1.1)
```

---

## 🔗 Links

- **GitHub Releases:** https://github.com/HenryBui21/VSBTek-Office-AIO/releases
- **Release v1.0 Notes:** [RELEASE_NOTES_v1.0.md](RELEASE_NOTES_v1.0.md)
- **Release v1.1 Notes:** [RELEASE_NOTES_v1.1.md](RELEASE_NOTES_v1.1.md)

---

## 💡 Tips

1. **Tag sẽ tự động tạo** khi bạn publish release
2. **v1.1 là latest** - User sẽ thấy đầu tiên
3. **v1.0 vẫn available** - Để tham khảo lịch sử
4. **Release notes được format đẹp** - GitHub hỗ trợ Markdown

---

**Note:** Tags chỉ được tạo KHI publish release, không tạo trước.

