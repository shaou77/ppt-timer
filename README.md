# PPT 演讲计时器

一个用于 Windows PowerPoint 放映的小型半透明悬浮计时器。

## 功能

- 自动检测 PowerPoint 放映开始，并自动启动倒计时
- 左侧显示倒计时，右侧显示 `剩余张数/总张数`
- 时间到 0 后继续负计时，并闪烁提醒
- 设置页支持自定义时长，以及 `10 / 20 / 30 分钟`快捷按钮
- 右上角提供设置、关于、关闭按钮
- 默认显示在屏幕右上角，可拖动位置
- 便携版运行，不需要安装

## 使用

1. 从 Release 下载 `PPT-Timer-Portable.zip`
2. 解压后双击 `PPT-Timer.exe`
3. 打开 PowerPoint 并开始放映
4. 需要调整时间时，点击右上角齿轮

设置会保存在程序同目录的 `PPT倒计时设置.ini`。

## 兼容性

- Windows 10 / 11
- Microsoft PowerPoint 桌面版

## 开发

本项目使用 C# WinForms 编写，可用仓库中的 `build.ps1` 构建。

```powershell
.\build.ps1
.\dist\PPT-Timer.exe --self-test
```

## 关于

- 版本：1.1
- 作者：Nick Lee
- 邮箱：shaou77@sina.com
