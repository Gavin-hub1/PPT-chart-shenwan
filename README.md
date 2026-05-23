# PowerPoint 图表格式化加载项

这个项目生成一个名为“申万研报格式”的 PowerPoint VBA 加载项，用于一键格式化选中的图表。

## 功能

- 横坐标轴：主刻度线外侧，轴线与标签为参考图蓝色，轴线粗细 1 磅。
- 左右纵坐标轴：主刻度线外侧，无次刻度线，轴线与标签为参考图蓝色，轴线粗细 1 磅。
- 删除主要/次要网格线。
- 图例放在下方。
- 删除图表标题。
- 柱形/条形系列填充参考图蓝色。
- 折线系列线条为橙色。
- 只调整已有数据标签的字体和位置，不新增数据标签，不改数字格式。
- 不改变系列所属坐标轴，不改变原始数据，不改变坐标轴数字格式。

默认颜色：

- 参考图蓝色：`RGB(37, 24, 180)`
- 橙色：`RGB(237, 125, 49)`

## 生成加载项

在 Windows 且已安装 PowerPoint 的环境中运行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\Build-ChartFormatterAddin.ps1
```

生成文件位于 `dist`：

- `ChartFormatter.pptm`：可打开调试的宏演示文稿。
- `ChartFormatter.ppam`：可安装的 PowerPoint 加载项。

如果脚本提示 PowerPoint 阻止 VBA 导入，请在 PowerPoint 中启用：

`文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置 > 信任对 VBA 项目对象模型的访问`

如果在自动化/沙箱会话里看到 `0x80070520` 或无法创建 `PowerPoint.Application`，请在已登录桌面的普通 PowerShell 窗口中运行同一条构建命令。

## 安装

1. 打开 PowerPoint。
2. 进入 `文件 > 选项 > 加载项`。
3. 在底部 `管理` 中选择 `PowerPoint 加载项`，点击 `转到`。
4. 点击 `添加新加载项`，选择 `dist\ChartFormatter.ppam`。
5. 启用加载项后，功能区会出现 `申万研报格式 > 图表格式 > 申万研报格式`。

## 使用

1. 在幻灯片中选中一个或多个图表。
2. 点击 `申万研报格式 > 申万研报格式`。
3. 插件会格式化选中图表，并跳过非图表对象。

## 手动导入宏

如果暂时不生成加载项，也可以打开 PowerPoint VBA 编辑器，将 `src\ChartFormatter.bas` 导入到任意 `.pptm` 文件中，然后运行宏 `FormatSelectedCharts`。
