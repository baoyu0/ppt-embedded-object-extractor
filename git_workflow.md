# Git 工作流指南

## 分支策略

本项目采用双分支工作流：

- **main** - 主分支，用于发布稳定版本
- **develop** - 开发分支，用于日常开发和测试

## 开发流程

### 1. 切换到开发分支
```bash
git checkout develop
```

### 2. 拉取最新代码
```bash
git pull origin develop
```

### 3. 进行开发和修改
- 修改代码
- 添加新功能
- 修复bug
- 运行测试

### 4. 提交更改
```bash
# 添加修改的文件
git add .

# 提交更改（使用有意义的提交信息）
git commit -m "feat: 添加新功能描述" 
# 或
git commit -m "fix: 修复bug描述"
# 或
git commit -m "docs: 更新文档"
```

### 5. 推送到开发分支
```bash
git push origin develop
```

### 6. 测试验证
在develop分支上进行充分测试：
- 功能测试
- 单元测试
- 集成测试
- 性能测试

### 7. 合并到主分支
当开发分支的代码经过充分测试后，合并到main分支：

```bash
# 切换到main分支
git checkout main

# 拉取最新的main分支代码
git pull origin main

# 合并develop分支
git merge develop

# 推送到远程main分支
git push origin main
```

## 提交信息规范

使用以下前缀来标识提交类型：

- `feat:` - 新功能
- `fix:` - bug修复
- `docs:` - 文档更新
- `style:` - 代码格式化（不影响功能）
- `refactor:` - 代码重构
- `test:` - 添加或修改测试
- `chore:` - 构建过程或辅助工具的变动

## 分支保护

建议在GitHub上设置分支保护规则：

1. 保护main分支
2. 要求通过Pull Request才能合并
3. 要求代码审查
4. 要求状态检查通过

## 创建Pull Request

当需要将develop分支合并到main时，建议使用Pull Request：

1. 在GitHub上创建从develop到main的Pull Request
2. 添加详细的描述说明
3. 请求代码审查
4. 确保所有检查通过后再合并

## 快速命令参考

```bash
# 查看当前分支
git branch

# 查看所有分支（包括远程）
git branch -a

# 切换分支
git checkout <branch-name>

# 创建并切换到新分支
git checkout -b <new-branch-name>

# 查看分支状态
git status

# 查看提交历史
git log --oneline

# 查看分支差异
git diff main..develop
```

## 注意事项

1. **始终在develop分支进行开发**，不要直接在main分支修改
2. **定期同步**：经常从远程仓库拉取最新代码
3. **小步提交**：频繁提交小的更改，而不是一次性提交大量修改
4. **有意义的提交信息**：写清楚每次提交做了什么
5. **充分测试**：在合并到main之前确保代码经过充分测试
6. **代码审查**：重要更改建议通过Pull Request进行代码审查

## 紧急修复流程

如果main分支发现紧急bug需要修复：

```bash
# 从main分支创建hotfix分支
git checkout main
git checkout -b hotfix/urgent-fix

# 修复bug并测试
# ...

# 提交修复
git add .
git commit -m "hotfix: 修复紧急bug描述"

# 合并到main
git checkout main
git merge hotfix/urgent-fix
git push origin main

# 同时合并到develop
git checkout develop
git merge hotfix/urgent-fix
git push origin develop

# 删除hotfix分支
git branch -d hotfix/urgent-fix
```

这样的工作流确保了代码的稳定性和开发的灵活性！