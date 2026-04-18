---
name: push-code-to-github
overview: 将本地已提交的代码推送到GitHub远程仓库，解决之前的认证问题
todos:
  - id: update-gitignore
    content: 更新.gitignore，添加*.spec和query_history*.json排除规则
    status: completed
  - id: verify-git-status
    content: 确认git状态，确保所有源码文件正确暂存
    status: completed
    dependencies:
      - update-gitignore
  - id: github-auth
    content: 使用[skill:Browser Automation]完成GitHub Git Credential认证
    status: completed
    dependencies:
      - verify-git-status
  - id: push-code
    content: 执行git push -u origin master推送代码到远程仓库
    status: completed
    dependencies:
      - github-auth
  - id: verify-push
    content: 验证推送结果，确认远程仓库代码完整
    status: completed
    dependencies:
      - push-code
---

## Product Overview

将业务报表项目的所有代码提交到远程GitHub仓库 (https://github.com/wangkai789/codebudy.git)

## Core Features

- 将本地已commit但未推送的代码推送到GitHub远程仓库
- 之前多次尝试因token认证失败(403)，需要解决认证问题后完成推送
- 确保敏感文件（query_history.json）不被上传

## 当前状态

- 本地分支: master，最新提交 e74bea2
- 远程仓库: https://github.com/wangkai789/codebudy.git（空仓库）
- 问题：之前所有push尝试均返回403权限错误

## Tech Stack

- Git版本控制
- GitHub HTTPS认证（Personal Access Token）

## Implementation Approach

1. **更新.gitignore**：添加敏感文件排除规则（*.json历史记录文件、*.spec打包配置文件）
2. **确认本地代码状态**：确保所有源码文件已正确纳入git管理
3. **使用Browser Automation方式登录GitHub**：通过浏览器自动化完成Git Credential认证，避免token在命令行中被拒绝的问题
4. **执行git push**：将master分支推送到远程origin

## Implementation Notes

- .gitignore需要补充：`*.spec`、`query_history*.json` 排除规则
- 之前的403错误可能是token权限不足或token格式问题，改用浏览器交互式认证更可靠
- 推送前需确认remote URL正确指向 https://github.com/wangkai789/codebudy.git

## Agent Extensions

### Browser Automation / playwright-cli

- **Purpose**: 通过浏览器自动化方式完成GitHub Git Credential Manager的交互式登录认证，解决命令行token被拒绝(403)的问题
- **Expected outcome**: 成功配置Git凭据，使git push能够正常执行并完成代码推送