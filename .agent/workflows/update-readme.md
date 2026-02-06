---
description: 自动将新的commit特性提取到README更新日志中
---

# 更新 README 更新日志

此workflow用于自动将最近的commit信息提取并添加到README.md的更新日志部分。

## 步骤

### 1. 查看最近的commit记录

// turbo
```bash
git log --oneline -5
```

查看最近的5条commit记录，确认哪些是新的、需要添加到README的。

### 2. 获取要添加的commit详细信息

// turbo
```bash
git log --format="%h - %s (%ad)" --date=short -1
```

这会输出最新commit的格式化信息，格式为：`短哈希 - 提交信息 (日期)`

如果需要获取多条commit，可以修改 `-1` 为需要的数量，如 `-3` 获取最近3条。

### 3. 编辑 README.md

打开 `README.md` 文件，在 `### 📝 更新日志` 部分的第一个列表项之前，添加新的commit记录。

格式示例：
```markdown
* abc1234 - [feat] 新增某某功能 (2025-06-15)
```

**Commit信息规范建议**：
- `[feat]` - 新功能
- `[fix]` 或 `[bug]` - 修复bug
- `[refactor]` - 重构
- `[docs]` - 文档更新
- `[style]` - 样式调整

### 4. 验证更改

// turbo
```bash
git diff README.md
```

确认README的更改是正确的。

### 5. 提交更改（可选）

如果需要将README更新也作为单独的commit：

```bash
git add README.md
git commit --amend --no-edit
```

使用 `--amend` 将README更新合并到最后一次commit中，避免产生额外的commit。

---

## 💡 提示

- 如果commit信息本身已经足够清晰，可以直接复制使用
- 如果需要更详细的描述，可以手动编辑添加的内容
- 建议在每次发布新版本或添加重要功能后运行此workflow
