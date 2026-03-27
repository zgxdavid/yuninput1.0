# 匀码输入法功能与文档总索引

更新时间：2026-03-24

## 1. 功能总览

### 1.1 核心输入与候选
- 中文/英文模式切换
- 全角/半角模式切换
- 候选窗跟随与分页显示
- 数字键选词、空格上屏、回车上屏策略
- Tab/Shift+Tab 候选导航
- 连续上屏与余码保留

### 1.2 学习与个性化
- 用户置顶：Ctrl+1-9
- 候选屏蔽：Ctrl+Delete
- 词频学习持久化
- 上下文关联学习持久化
- 自动造词独立持久化
- 自造词管理入口（配置工具）

### 1.3 词库与模式
- zhengma-large（大字符集郑码）
- zhengma-large-pinyin（拼音单字输入，候选右侧显示郑码提示）
- `data/zhengma.dict` 保留为导入参考词表，不再作为配置工具中的独立运行模式
- 词库导入链路（Fcitx/IBus table -> .dict）

### 1.4 安装与运维
- 构建脚本
- 安装启用脚本（含提权）
- 注册/反注册脚本
- 卸载清理脚本
- 运行状态诊断脚本

## 2. 文档索引

### 2.1 项目入口
- README.md：项目总说明、构建与安装入口

### 2.2 需求与策略
- docs/requirements-audit.md：需求对账与完成度分级
- docs/zhengma-plan.md：郑码策略与导入规则说明

### 2.3 模式验证
- docs/mode-e2e-checklist.md：两模式验收流程
- docs/mode_samples.tsv：模式样本数据
- docs/multi-monitor-candidate-window-checklist.md：候选窗多显示器定位回归清单
- docs/multi-monitor-candidate-window-test-log-template.tsv：多显示器回归记录模板（含 Result/Notes/IssueID）
- docs/multi-monitor-candidate-window-test-log-template.csv：多显示器回归记录模板 CSV 版（便于 Excel 直接打开）

### 2.4 今日交接
- docs/handoff-20260324.md：本次收尾、清理、打包与现场保留信息

## 3. 关键脚本入口

- scripts/build_release.ps1：Release 构建
- scripts/build_gui_installer.ps1：GUI 安装包构建
- scripts/install_enable.ps1：本机安装启用
- scripts/uninstall_clean.ps1：卸载与注册表残留清理
- scripts/inspect_ime_state.ps1：IME 状态诊断

## 4. 现场保留原则

- 保留 diagnostics 下所有快照目录
- 保留 docs 下所有需求、验收、规划文档
- 删除可再生成的构建中间产物，不删除源码与诊断证据
