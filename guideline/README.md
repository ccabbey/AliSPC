# AliSPC 项目交接指南

>  written 2026-11-13，供项目迁移到新机器后接手的 agent / 开发者快速继承。
> 本目录（guideline/）为 UTF-8 Markdown，**不导入 VBE**；与 src/ 的 GBK 源文件互不影响。

## 项目是什么

在 **Excel 2016 VBA** 中复刻 **Minitab 的过程能力正态图**：直方图 + overall 正态曲线 + 均值/规格/_sigma 参考线 + 右上角统计框（N / Mean / StDev / Cp / Cpk / Pp / Ppk / Cpm）。

- 工作簿：`AliSPC_Dev2.0.0.261109.xlsm`（用户维护，**不是**代码基准）。
- 前身 Dev1.0（`AliSPC_Dev1.0.0.241104.xlsm`）已弃用：计算链存在但绘图链从未接通。
- 实现路线：XY 散点画轮廓/曲线/竖线 + ×10000 归一化面积系列叠加填充。这是被 Excel 对象模型逼出来的方案，不是自由选择，详见 `excel-rendering-notes.md`。
- 开发工具链：VBE + Rubberduck（注解与属性维护）。

## 目录约定

| 路径 | 说明 |
|---|---|
| `src/` | **唯一代码基准**。17 个平铺文件，GBK(GB2312) + CRLF，由用户从 VBE 导出。改代码只改这里。 |
| `*.xlsm` | 用户维护的工作簿。不要解析、不要修改、不要试图修复其属性。未保存的工作簿是旧快照。 |
| `todo.md` | 用户的需求清单（GBK 编码，读取前先解码）。需求以用户口头/此文件为准。 |
| `guideline/` | 本文档集（UTF-8）。 |

## 接手阅读顺序

1. 本文件（5 分钟定位）。
2. `working-agreement.md` —— **协作铁律，动手前必读**（落盘权限、GBK 机制、验证审批、token 节约）。
3. `status-and-tasks.md` —— 现状、已锁定决策、待办、未验证风险、接手第一步。
4. `architecture-and-style.md` —— 分层契约、模块清单、VBA 编码约定与编译级教训（改代码前读）。
5. `excel-rendering-notes.md` —— 已探明的 Excel 轴拓扑与渲染机制（动渲染层前读，避免重蹈覆辙）。

## 铁律速览（详见 working-agreement.md）

1. **默认不落盘**：只在聊天里给代码/方案，用户自己更新并重新导出平铺文件。用户明确说"落盘"才动文件。
2. **src 必须保持 GBK + CRLF**：Edit/Write 工具拒绝非 UTF-8 文件；落盘走 python 脚本机制（`io.open(..., encoding="gbk", newline="")` + 锚点断言 + 回读校验）。
3. **改前重读磁盘**：用户会并发改代码，任何缓存快照都不可信。
4. **验证需审批**：测试/PowerShell/COM/后台 Excel 等工具链，先整理成待办、获用户明确同意才用。python 作为 GBK 编辑机制（含回读断言）是已接受的例外。
5. **不要过度设计**：接口先行、计算与渲染解耦、不为假设需求加配置项。
6. **分工**：PredeclaredId / Rubberduck / 工作簿侧属性 / 编译与渲染验证归用户；agent 不碰。

## 当前状态一句话

核心绘图管线已完成；2026-11-13 落盘四项（全 bin 边界标签、删除水平对齐逻辑、纵轴固定贴左、数据源表头→图表标题）。**等待用户侧编译 + 渲染回归**（清单见 status-and-tasks.md）；todo7（PlotMeta 参数持久化）挂起待用户确认是否仍需要。