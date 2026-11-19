# 现状与待办（截至 2026-11-13）

## 里程碑

| 日期 | 事件 |
|---|---|
| 2024-11 | Dev1.0（241104）弃用：计算链存在、绘图链未接通、无 UserForm。 |
| 2026-11-09 | Dev2.0（261109）重启，后端重新设计（分层 + Spec 驱动）。 |
| 2026-11-10 | 经批准的 COM 探针确定轴拓扑（见 excel-rendering-notes.md）；枚举作用域/保留字编译问题清理。 |
| 2026-11-11 | todo2（规格全可选）/ todo5（密度模式）/ todo3（SeriesDump）落盘。 |
| 2026-11-12 | todo1（bin 网格刻度）/ todo4（axisLabels 接管）/ Decimals 显示精度落盘。 |
| 2026-11-13 | FormatNumber 合并（用户应用）；四项落盘：全 bin 边界标签(#23)、删除水平对齐逻辑(#24)、纵轴固定贴左(#25)、表头→标题(#26)。 |

## 已实现能力

- **数据摄入**：Range（late-bound，单区域）/ 一维数组 / 单行或单列二维数组；非数值跳过；表头（hasHeader 跳首元素取文本 / header 显式传入）；N≥2 校验。
- **统计**：Mean、StdevOverall(N-1)、StdevWithin(MRbar/1.128)、min/max；Cp/Cpk/Pp/Ppk/Cpm（无规格→N/A）。
- **分箱**：Sturges / Sqrt / FixedCount / FixedWidth，等宽，MAX_BINS 上限，末边界钉死。
- **绘图**：直方图填充+轮廓、overall 正态曲线（平滑）、均值线、LSL/USL/Target、±k/2·sigma 线、轴标签接管（全 bin 边界 + 彩色辅助值）、统计框、密度/频次双模式、Decimals 显示精度、bin 网格刻度、纵轴贴左、标题（opt.Title > 表头回退）。
- **导出**：SeriesDump.WriteData 按系列导出 X/Y 到工作表（列 stride 3）。
- **测试**：CheckCapabilityDemo（60 硬编码点，spec 9/11/10，binFixedCount 10，EXPECTED_SERIES=10）；DrawCapabilityDemo（Sheet3!A1:A999，hasHeader=True，Decimals=2，binFixedCount 20，密度模式，无规格线，标题走表头）。

## 已锁定决策（不要重开）

1. 只画 **overall** 正态曲线（不画 Minitab 的 within/overall 双曲线）。
2. **无超规格红色 shading**，填充单系列。
3. **等宽分箱**；×10000 面积映射方案接受。
4. 密度纵轴显示**圆整到 Decimals**（用户推翻 General 推荐）。
5. Cpm = within-sigma / Taguchi（用户对照 Minitab 确认关闭）。
6. 轴标签路线 = **隐藏散点 DataLabel**（用户选定）。
7. 水平对齐需求**多余，已删除**（2026-11-13，用户判定）。
8. 纵轴**永远贴左**（恒定行为，不做配置项）。

## 待办

- **#21 用户侧回归**（不经过工具链，接手后先问结果）：
  1. 编译通过（AnnotationSpec.Create/Initialize 签名已变，确认无遗漏调用点）。
  2. CheckCapabilityDemo PASS=10。
  3. DrawCapabilityDemo：20 bin 全边界标签在 Decimals=2 下是否重叠；原生刻度隐藏；彩色 DataLabel 在轴下方（**重点：线+标记隐藏时标签是否可见**）。
  4. 统计框右缘位置不变、框内段落变左对齐。
  5. 负 X 数据纵轴贴左。
  6. Sheet3!A1 为表头时标题显示表头、N 不含表头（**若 A1 是数值，去掉 demo 里的 True**）。
  7. Decimals 0/5、ShowAxisLabels=False 回退原生标签、dump 含 axisLabels 块。
- **#20 todo7：PlotMeta 参数持久化**（旧版 todo 遗留；新版 todo.md 未再提及，**接手后先与用户确认是否仍需要**，细节需用户补充）。
- **#13（冻结）**：可选的工具链预验证，仅在用户明确同意时启用。
- `todo.md` 是用户的活需求清单（GBK，先解码再读）；2026-11-13 版的 3 项（对齐删除/纵轴贴左/表头）已全部完成。

## 未验证风险（全部需用户侧确认）

1. **DataLabel 在线+标记隐藏时的可见性**（todo4 最大风险，机制上未证实）。
2. ";;;" 隐藏原生标签且保留占位。
3. 刻度恰好落在 bin 边界（吸附 + MajorUnit 的渲染效果）。
4. Crosses=xlAxisCrossesMinimum 在负 X 数据下的实际表现。
5. demo 假设 Sheet3!A1 是表头（若是数值会吞掉第一个数据点、标题变数字）。

## 已知怪癖（设计使然，非 bug）

- 辅助标签落在 bin 边界 eps(2% span) 内会被丢弃（边界优先）；数值几乎相同，可接受。
- BinCount>50 时边界仍全显示（eps=0），但视觉重叠不可避免 → 减 Decimals/加宽/减 bin。
- 统计框段落左对齐（2026-11-13 起），框位置不变。
- SeriesDump 的 Err.Raise source 字符串可能仍写 "SeriesDump.Write"（装饰性，已提示用户，可能已修）。
- 复合图 Chart.ChartType=-4111 无名值，永远不对它分支。

## 本机产物（不随项目迁移，结论已固化在本目录）

- 旧机器 `Temp/alispc_gen/`：GBK 编辑脚本、axes_probe*.ps1 探针。
- 旧机器 agent 记忆：`C:/Users/li1a/.qoder-cn/projects/D---CSProjects-vba/memory/`（本 guideline 是其可移植继任者）。
- 新机器以 `src/` 平铺文件 + `todo.md` + 本目录为唯一基准。

## 接手第一步

1. 读 `working-agreement.md`（铁律）。
2. 解码读 `todo.md`，看有无新需求。
3. **重读**相关 src 文件（不信任任何缓存快照；用户可能又改过）。
4. 向用户确认：#21 回归结果、todo7 去留、Sheet3!A1 表头假设。
5. 默认只给聊天代码；用户说"落盘"才用 python GBK 机制写文件。