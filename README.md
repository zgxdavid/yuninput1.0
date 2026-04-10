# 匀码输入法（yuninput）

English README: README.en.md

匀码输入法（yuninput）是一个面向 Windows 11 TSF（Text Services Framework）的中文输入法项目，当前重点是把郑码输入体验、学习能力、配置工具和安装打包链路做成一套能持续迭代的工程化实现。

它现在已经不只是一个 TSF 骨架，而是一套可构建、可清理、可打包、可学习、可调频、可管理用户词条和自动造词的 Windows 输入法工程。

## 项目状态

- 当前阶段：持续开发中，但已经具备日常测试和持续回归的基础能力
- 目标平台：Windows 11 x64
- 技术路线：C++ TSF Text Service + Win32 候选窗 + PowerShell 安装/诊断脚本 + C# 配置工具
- 默认词库模式：zhengma-all
- 当前定位：面向郑码/形码输入实验与工程实现，不宣称为成熟商业成品

## 当前能力

### 输入与候选

- 支持中文/英文模式切换、全角/半角切换、候选窗跟随光标、屏幕边界约束和基础状态菜单
- 支持数字键 1-9 选词、空格上屏、回车优先精确候选、Esc 清空、Backspace 回退
- 支持 Tab/Shift+Tab 导航候选，支持 [ ]、, .、PgUp、PgDn 等翻页
- 支持 Ctrl+1-9 按当前编码置顶候选，支持 Ctrl+Delete 屏蔽当前候选
- 候选窗支持单字郑码提示、页脚提示和模式切换后的即时刷新

### 词库模式与编码数据

- 默认加载 zhengma-all.dict，兼顾常用郑码输入与较合理的默认词序
- 支持 zhengma-large-pinyin 模式，用于拼音单字输入并附带郑码提示
- 支持通过配置工具或运行时热键即时切换词库模式并重载用户学习数据
- 支持从外部 Fcitx / IBus Table 风格源表导入 .dict 词库
- 支持 zhengma-single.dict 作为自动造词与郑码提示的单字编码来源

### 排序、调频与学习

- 支持上屏调频，常用单字和词语会真正前移，而不只是高亮变化
- 支持按编码维度的用户置顶和屏蔽，并持久化到用户数据文件
- 支持上下文联想学习与持久化，可在配置工具中查看和清理
- 支持把正式标准郑码词条保持在较靠前位置，同时让高频真实使用项持续争先
- 排序逻辑已调整为同时考虑用户置顶、学习频次、自动造词/用户词条、静态词频和加载顺序

### 自动造词与用户词典

- 支持连续上屏片段的自动造词和会话期短语缓存
- 会话中生成的短语在首次被确认后可晋升为持久化词条，写入 yuninput_user.dict
- 用户词典支持保留 auto 标签，自动造词不会在重载或切窗口后丢失
- 配置工具可直接管理 yuninput_user.dict 中的手工词、自动词和屏蔽项
- 提供造词集中审阅面板，可按来源查看近期进入词典的短语

### 配置、诊断与运维

- 提供 yuninput_config.exe 图形配置工具
- 支持设置候选页大小、自动顶屏最短码长、中文标点、智能成对符号、空候选提示音、Tab 导航、回车策略、上下文联想上限等
- 支持语言栏菜单和状态菜单打开配置工具、系统输入设置和运行日志
- 提供 inspect_ime_state.ps1、uninstall_clean.ps1、clean_delete_ime.ps1 等诊断/清理脚本
- 支持当前用户修复 HKCU profile、刷新 ctfmon、检查 DLL 注册路径与词库加载状态

### 构建与发布

- 支持 CMake + MSVC 全量构建 yuninput.dll、sort_probe、user_dict_builder 等 C++ 产物
- 支持单独构建 C# 配置工具 yuninput_config.exe
- 支持 WiX v4 生成 MSI，产物默认输出到仓库根目录的 Yuninput1.2.msi
- MSI 已打包主 DLL、配置工具、安装/卸载脚本、默认词库、用户词库模板和规则文件

## 快速开始

### 构建要求

- Windows 11
- Visual Studio 2022（含 C++ 桌面开发）
- CMake 3.21+
- WiX CLI（打 MSI 时需要）
- PowerShell 5.1 或更新版本

### 本地全量构建

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/build_release.ps1
./scripts/build_config_app.ps1
```

这会生成：

- build/Release/yuninput.dll
- build/Release/yuninput_sort_probe.exe
- build/Release/yuninput_user_dict_builder.exe
- tools/yuninput_config.exe

### 生成 MSI

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/build_msi_wrapper.ps1 -Version 1.2.0 -OutputName Yuninput1.2.msi -SkipDictionaryGeneration
```

默认输出：

- ./Yuninput1.2.msi

说明：当前 scripts/generate_user_dict.ps1 在 Extend User Dictionary Variants 阶段仍可能卡住，因此现阶段打包通常使用 -SkipDictionaryGeneration，直接复用仓库内已有词库产物。

### 清理旧安装

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/clean_delete_ime.ps1 -SkipElevation
```

如需更彻底地清理机器级注册残留，请在管理员 PowerShell 中运行：

```powershell
./scripts/uninstall_clean.ps1
```

## 日常使用与热键

- Ctrl+F2 可直接打开配置工具；Ctrl+Shift+F2 也会命中同一入口
- F2 或 APPS 可打开状态菜单
- Ctrl+Shift+F3 切到 zhengma-large-pinyin 模式
- Ctrl+Shift+F4 切回 zhengma-all 模式
- Ctrl+Shift+M 可将当前选中候选提升为手工词条
- Ctrl+1-9 置顶当前候选；Ctrl+Delete 屏蔽当前候选
- [ ]、, .、PgUp、PgDn 可翻页；Tab/Shift+Tab 可切换高亮候选

状态菜单与语言栏按钮当前可直接打开：

- 配置工具
- System Input Settings
- 运行日志

排查输入法注册状态：

```powershell
./scripts/inspect_ime_state.ps1
```

如果报告里出现 HKLM profile exists: True 但 HKCU profile exists: False，可补当前用户侧 profile 并重启 ctfmon：

```powershell
./scripts/inspect_ime_state.ps1 -RepairCurrentUserProfile -RestartCtfmon
```

完整诊断结果会写入 %LOCALAPPDATA%\yuninput\ime_state_report.txt。

## 排序与学习逻辑

当前候选排序不是单一的静态词频排序，而是把静态词库质量和真实使用行为叠加起来：

1. 先看是否被用户显式置顶或屏蔽。
2. 再看该编码下的真实上屏学习结果，包括单字和词语的累计使用频次。
3. 再看是否属于用户词条或已晋升的自动造词条目。
4. 最后回退到静态词库分数、字集优先级和原始加载顺序。

这一轮实现特别修正了两个过去容易错位的问题：

- 学习后的候选现在会真实移动到前面，而不是只改变默认选中高亮。
- 会话中产生的自动造词在首次确认后会进入用户词典，切窗口、重载词库后仍能保留。

目标是保持两个约束同时成立：

- 正式标准郑码编码词条在初始状态下仍然排在相对前面。
- 真正被反复输入的单字和词语，应该能尽快冲到更靠前的位置。

## 词库与数据文件

当前工程的核心数据文件包括：

- data/zhengma-all.dict：默认主词库
- data/zhengma-pinyin.dict：拼音单字模式词库
- data/zhengma-single.dict：单字郑码来源，用于提示和造词构码
- data/yuninput_user.dict：用户词典，包含手工词和可持久化的 auto 标签词条
- data/yuninput_user-extend.dict：随安装包分发的扩展用户词条集
- data/*.rules：词组构词和元数据规则文件

导入外部码表：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/import_table_dict.ps1 -SourcePath C:\temp\zhengma-large.txt -OutputPath ./data/zhengma-large.dict
```

生成拼音单字词库：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/generate_pinyin_single_char_dict.ps1
```

两字词构码回归可使用：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./scripts/run_twochar_phrase_regression.ps1
```

## 性能优化进展

围绕“大词库 + 实时选词 + 自动造词”这个场景，当前已经落地过一批直接影响流畅度的优化：

- 候选查询引入前缀范围索引，减少整表扫描
- 增加查询缓存，降低重复编码命中的计算成本
- 候选排序优先做有限范围处理，避免每次都对全部候选做重排
- 用户数据落盘改为延迟/批量写入，避免频繁同步 I/O 卡住提交路径
- 新增单条词条的增量索引插入逻辑，减少新增用户词或自动词时的全量 RebuildIndex 开销
- 保留单字郑码提示缓存，减少运行时重复解析单字码表

这意味着当前的主要热点已经从“每次上屏都同步写文件”转向“首次晋升新词时的引擎内维护成本”。最近一次修改已经把单条新增词条从全量索引重建改成增量插入，用来继续压低首次自造词确认时的迟滞感。

## 仓库结构

- include/：头文件
- src/：输入引擎、TSF 入口、候选窗与学习逻辑
- scripts/：构建、安装、卸载、诊断、回归和打包脚本
- tools/：配置工具、MSI 封装与辅助程序
- data/：词库、用户词典模板和规则文件
- docs/：设计说明、回归清单、排障记录和交接文档
- diagnostics/：现场快照、日志与安装/回归留档
- third_party/：第三方源表或上游快照

## 已知边界

- 项目仍处于持续开发阶段，词库、排序和 UI 细节还会继续迭代
- scripts/generate_user_dict.ps1 的扩展词典生成阶段仍有卡住问题，尚未彻底修复
- 机器级 HKLM 清理仍可能需要管理员权限
- 词库覆盖率和静态默认排序质量仍取决于导入源表质量
- 某些第三方词库或派生产物能否公开再分发，取决于其上游许可状态

## 路线方向

1. 继续优化首次晋升新词、自造词提交和候选刷新路径的延迟
2. 继续增强郑码词组规则、上下文联想质量和用户数据管理体验
3. 完善更稳定的词典生成链路，消除当前 extend 阶段的卡住问题
4. 继续补齐发布、升级、卸载和多环境回归流程

## 开源许可与声明

- 本项目采用 GPL-3.0-or-later 开源许可，见 LICENSE
- 中文许可说明见 LICENSE.zh-CN.md（法律约束以 LICENSE 原文为准）
- 项目原创代码权属声明见 NOTICE.md 与 NOTICE.zh-CN.md
- 第三方来源与分发注意事项见 THIRD_PARTY_NOTICES.md 与 THIRD_PARTY_NOTICES.zh-CN.md
- third_party/fcitx/ 下保存的是 Fcitx 格式源表快照；在公开发布前，应补齐准确上游链接与许可文本

## 发布前资料

- 开源发布前检查清单：OPEN_SOURCE_RELEASE_CHECKLIST.md 与 OPEN_SOURCE_RELEASE_CHECKLIST.zh-CN.md
- GitHub 开源发布说明模板：GITHUB_RELEASE_TEMPLATE.md 与 GITHUB_RELEASE_TEMPLATE.zh-CN.md