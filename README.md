# 匀码输入法（yuninput）

[![License](https://img.shields.io/github/license/zgxdavid/yuninput1.0)](LICENSE)
[![Last Commit](https://img.shields.io/github/last-commit/zgxdavid/yuninput1.0)](https://github.com/zgxdavid/yuninput1.0/commits/main)

- GitHub 仓库：https://github.com/zgxdavid/yuninput1.0
- 许可证：GPL-3.0（项目采用 GPL-3.0-or-later 分发策略，详见 `NOTICE.md`）
- 发布页：https://github.com/zgxdavid/yuninput1.0/releases

匀码输入法（yuninput）是一个面向 Windows 11 TSF（Text Services Framework）的中文输入法项目。

当前仓库重点提供一套可构建、可安装、可扩展的 Windows 输入法骨架，包括输入引擎、候选窗、配置工具、安装脚本以及 EXE/MSI 打包链路，并逐步补齐类似郑码输入体验所需的词库导入、学习能力和词组规则支持。

## 项目状态

- 当前阶段：开发中
- 目标平台：Windows 11 x64
- 技术路线：C++ TSF Text Service + PowerShell 安装脚本 + Windows 配置工具
- 当前定位：工程化输入法框架，不是完整商业成品

## 当前能力

- C++ DLL 工程，支持 CMake + MSVC 构建
- TSF Text Service 基础 COM 框架
- `DllRegisterServer` / `DllUnregisterServer`
- 候选窗显示、跟随光标、屏幕边界约束
- 数字键 `1-9` 选词
- 候选置顶：`Ctrl+1-9`
- 删词屏蔽：`Ctrl+Delete`
- 候选翻页：`-` / `=`，兼容 `[` / `]`、`PgUp` / `PgDn`
- Tab/Shift+Tab 候选导航
- 空格上屏、回车优先精确候选、Esc 清空、Backspace 回退
- 连续上屏与余码保留
- 上屏调频与持久化
- 上下文学习与持久化
- 配置工具支持用户词条、自动造词、上下文学习管理
- 自动造词独立存储，不混入手工词表
- 第一版郑码词组构词规则已接入自动造词
- 可导入外部 Fcitx / IBus Table 风格码表
- 支持生成 EXE 安装包与 MSI 封装安装包

## 适合谁使用

- 想研究 Windows TSF 输入法实现的人
- 想在 Windows 上实验形码/郑码类输入体验的人
- 想基于现有工程继续扩展词库、学习策略和 UI 的开发者

如果你的目标是直接获得一个已经成熟、长期稳定、词库完备的最终输入法产品，这个仓库当前还不是那个状态。

## 快速开始

### 构建要求

- Windows 11
- Visual Studio 2022（含 C++ 桌面开发）
- CMake 3.21+

### 本地构建

```powershell
cmake -S . -B build -A x64
cmake --build build --config Release
```

### 一键构建并安装

管理员 PowerShell：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/install_enable.ps1 -Build
```

### 卸载

```powershell
./scripts/unregister_ime.ps1
```

如需清理安装残留并打开系统设置页：

```powershell
./scripts/uninstall_clean.ps1
```

## 配置与日常使用

- 在匀码输入法激活时，按 `Ctrl+F2` 可直接打开配置界面
- `F2` / `APPS` 可打开状态菜单
- 配置工具可管理用户词条、自动造词、上下文学习、候选策略与模式设置

排查输入法注册状态：

```powershell
./scripts/inspect_ime_state.ps1
```

如果报告里出现 `HKLM profile exists: True` 但 `HKCU profile exists: False`，可补当前用户侧 profile 并重启 `ctfmon`：

```powershell
./scripts/inspect_ime_state.ps1 -RepairCurrentUserProfile -RestartCtfmon
```

完整诊断结果会写入 `%LOCALAPPDATA%\\yuninput\\ime_state_report.txt`。

## 词库与郑码导入

当前工程支持将外部码表导入为 yuninput 可直接加载的 `.dict` 文件。

导入脚本：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/import_table_dict.ps1 -SourcePath C:\temp\zhengma-large.txt -OutputPath ./data/zhengma-large.dict
```

拼音单字模式词库可从 Fcitx 风格源表中提取：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/generate_pinyin_single_char_dict.ps1
```

重新安装时，`data/` 目录下的 `.dict` 文件会被复制到安装目录。

更多背景与策略说明见：

- `docs/zhengma-plan.md`

## 仓库结构

- `include/`：头文件
- `src/`：核心实现
- `scripts/`：构建、注册、安装、卸载与打包脚本
- `tools/`：配置工具与附属打包工具
- `data/`：词库与示例数据
- `third_party/`：第三方源表或上游快照
- `docs/`：功能说明、计划、审计与交接文档

## 安装包

当前仓库可生成以下发布产物：

- EXE 安装包：`yuninput_setup.exe`
- MSI 封装安装包：`yuninput_setup.msi`

MSI 当前采用单包封装方式，在安装阶段调用内置安装流程完成部署与注册。

## 已知边界

- 本项目仍在开发中，行为和配置方式可能继续调整
- 词库质量与覆盖率取决于导入源表质量
- 某些第三方词库或其派生产物是否可公开再分发，取决于其上游许可状态
- 如果某些词库不能附带发布，使用者需要自行导入

## 路线方向

1. 引入许可清晰、可分发的基础与扩展词库
2. 完善联想、造词、删词、置顶、符号策略与模式切换
3. 增强 TSF 组合串与系统级 UIElement 集成
4. 完善安装器签名、升级卸载与兼容性测试矩阵

## 开源许可与声明

- 本项目采用 `GPL-3.0-or-later` 开源许可，见 `LICENSE`
- 项目原创代码权属声明见 `NOTICE.md` 与 `NOTICE.zh-CN.md`
- 第三方来源与分发注意事项见 `THIRD_PARTY_NOTICES.md` 与 `THIRD_PARTY_NOTICES.zh-CN.md`
- `third_party/fcitx/` 下保存的是 Fcitx 格式源表快照；在公开发布前，应补齐准确上游链接与许可文本
- 关于郑码相关名称与知识产权的权利归属声明，见 `NOTICE.md`、`NOTICE.zh-CN.md` 与 `THIRD_PARTY_NOTICES.md`

## 发布前资料

- 开源发布前检查清单：`OPEN_SOURCE_RELEASE_CHECKLIST.md` 与 `OPEN_SOURCE_RELEASE_CHECKLIST.zh-CN.md`
- GitHub 开源发布说明模板：`GITHUB_RELEASE_TEMPLATE.md` 与 `GITHUB_RELEASE_TEMPLATE.zh-CN.md`
