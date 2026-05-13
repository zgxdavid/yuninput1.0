# 仓库审阅总结（2026-05-13）

## 范围与说明

- 审阅范围：`yuninput1.0` 目录及主要子目录（`src`、`include`、`scripts`、`tools`、`docs`、`diagnostics`、`data`）。
- 本文区分两类结论：
  1. 代码与脚本直接可验证结论
  2. 交接与诊断文档给出的阶段性结论

## 最近实现的功能（主线）

1. 输入体验与候选排序能力持续增强
   - 候选排序叠加了用户置顶、学习频次、上下文联想、自动造词、静态词频等多维策略。
   - 运行时支持词库模式切换与候选页策略。

2. 自动造词与用户词条管理能力增强
   - Session 自动造词有持久化状态、临时词元数据和晋升路径。
   - 自动词可晋升到用户词典，并与配置工具管理路径打通。

3. 运维与打包链路工程化
   - 构建、配置工具、MSI、GUI 安装器、诊断与清理脚本链路完善。
   - 新增 VS C++ 工作负载修复脚本，提升新机器环境自愈能力。

## 最近修复的问题（已落地）

1. ESC 键误吞问题修复（全屏/演示场景）
2. 自动造词回看边界与保留策略修复
3. 两字词短码联想增强
4. 四码通配查询规则调整（专用符号改为 `0`，限制位次）
5. 调频落盘时序问题修复（停用输入法时同步等待关键数据写盘）
6. session 自动造词保留策略修复（改为严格按最近 2000 汉字窗口保留）
7. 两码简码单字调频回退修复（兼容来源下恢复编码维度调频）

## 仍未完全解决的优化与风险

1. 郑码规则层能力仍非完整实现
   - 现阶段仍以表驱动为主，完整拆根与构词规则层仍有缺口。

2. 词典生成链路有已知卡点
   - `scripts/generate_user_dict.ps1` 在扩展阶段仍可能卡住，打包通常通过 `-SkipDictionaryGeneration` 规避。

3. 默认 VS 生成器链路在新机器上稳定性不足
   - 环境修复后，替代链路可构建，但默认 VS 生成器仍存在不稳定现象，需继续排障。

4. 系统集成与回归维度仍需持续验证
   - 任务栏名称/图标、候选窗多显示器行为等仍需持续回归。

## 构建与发布链路状态（截至 2026-05-13）

1. 核心二进制链路：可构建
2. 配置工具：可构建
3. MSI：可打包
4. GUI 安装器：可生成
5. 默认 VS 生成器链路：仍需继续稳定化

## 关键证据文件（抽样）

- 核心实现：
  - `src/TextService.cpp`
  - `src/CompositionEngine.cpp`
  - `src/CandidateWindow.cpp`
- 构建与打包：
  - `scripts/build_release.ps1`
  - `scripts/repair_vs_cpp_workload.ps1`
  - `scripts/build_msi_wrapper.ps1`
  - `scripts/build_gui_installer.ps1`
- 需求与规划：
  - `README.md`
  - `docs/requirements-audit.md`
  - `docs/zhengma-plan.md`
- 交接与诊断：
  - `diagnostics/handoff_20260511_143000.md`
  - `diagnostics/handoff_20260512_shutdown_save.md`
  - `diagnostics/latest_handoff.md`
