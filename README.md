# SlideSCI for Mac

这是 SlideSCI 的 Mac PowerPoint Office.js 版本。它是一个独立工程，所有代码都放在 `slidesci_for_mac/` 下，不引用 Windows VSTO 工程。

## 开发

```bash
npm install
npm run manifest:dev
npm run dev
```

默认开发地址是 `https://localhost:3000`。第一次运行时，先在 Safari 打开 `https://localhost:3000` 并信任本地开发证书，否则 PowerPoint 可能无法打开任务窗格。

开发模式下，如果要使用 Markdown 行内公式、块级公式或“插入 LaTeX 原生公式”，另开一个终端运行：

```bash
npm run helper
```

helper 监听 `http://127.0.0.1:17926`，负责通过 PowerPoint 自动化创建原生公式文本框。首次使用时 macOS 可能要求允许终端或 Node 自动化控制 Microsoft PowerPoint；如果拒绝，含公式模块会失败并在状态栏显示原因，不会静默降级成图片。

## 安装给最终用户

推荐安装方式：

```bash
npm run build
npm run install:mac
```

安装脚本会完成这些动作：

1. 编译本地 `SlideSCICompanion` watcher。
2. 复制 helper 到 `~/Library/Application Support/SlideSCI/`。
3. 注册 `launchd` 用户级 LaunchAgent，让 companion 常驻监听 PowerPoint 进程。
4. 渲染并复制 `manifest.xml` 到 PowerPoint 的 `wef` 侧载目录。

安装完成后，完全退出并重启 PowerPoint。此后：

- PowerPoint 打开时，companion 会自动拉起本地 helper。
- PowerPoint 关闭后，helper 会在短暂宽限期后自动退出。
- `SlideSCI` 入口位于“视图”选项卡右侧。

如需卸载：

```bash
npm run uninstall:mac
```

当前安装脚本仍要求本机存在 `node` 以运行 helper；仓库里已经加入了 companion/installer 基础设施，后续可继续把 helper 打包成独立可执行文件，去掉这个依赖。

Mac PowerPoint 本地侧载传统 XML manifest：

```text
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
```

把 `manifest.xml` 放入该目录后完全退出并重启 PowerPoint。加载成功后，会在“视图”选项卡出现 `SlideSCI` 分组和“打开 SlideSCI”按钮；也可以从“加载项”入口查找 `SlideSCI for Mac`。

如果加载项不出现：

1. 确认目录是 `/Users/<你的用户名>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`，不是 iCloud 或项目目录里的 `wef`。
2. 确认复制的是 `manifest.xml` 文件本身，不是 `slidesci_for_mac` 文件夹。
3. 删除旧 manifest 后重新复制新版 `manifest.xml`，再完全退出 PowerPoint。
4. 清理 Office 缓存后再试：删除 `~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches` 下的 Office 相关缓存。
5. 确认 `npm run dev` 正在运行，并且 Safari 可以打开 `https://localhost:3000/index.html`。

## 当前功能

- 图片自动排列：列最大宽度、统一高度、统一宽度瀑布流。
- 图片标题：为选中对象添加上标题或下标题。
- 图片标签：支持字母、数字、罗马数字、圆圈数字、中文数字模板。
- 内容插入：Prism 高亮代码块、Markdown 富文本/原生表格/原生公式/引用块、LaTeX 原生公式、LaTeX 图片。
- 格式工具：复制和粘贴中心位置、宽度、高度、文字/填充/边框基础格式。

## Markdown 与 LaTeX 验收

`test.md` 是固定回归样例。把它全文复制到 Markdown 输入框后，应按原文顺序插入标题、普通段落、普通列表、任务列表、表格、数学公式、代码块、引述块和有序列表，直到“归档步骤”结束。任务列表会转换成 `☑` / `☐` 前缀；表格会去掉单元格内的 Markdown 标记；行内公式会保留在同一个文本框中，并由 helper 直接创建为 PowerPoint 原生公式。

表格、代码块、块级公式和引述块会作为独立模块插入。某个模块在当前 PowerPoint API 下失败时，插件会继续插入后续模块，并在状态栏显示失败原因。表格优先使用原生 PowerPoint table：先创建空表格，只传 `left/top/width/height`，再通过 table cell API 逐格写入内容；如果该路径失败，才尝试 `values` 参数和文本框网格降级。

块级公式优先通过 helper 生成原生公式；helper 未运行或权限不足时，该公式模块会明确失败，防止把原生公式需求误做成图片。“插入 LaTeX 原生公式”只走 helper；“插入 LaTeX 图片”才使用 MathJax PNG 图片路径。图片插入优先使用矩形图片填充，避免依赖 Mac 端可能不可用的 `addPicture` preview API；如果图片路径全部不可用，才降级为 LaTeX 源码文本框。选中由本插件插入的公式后点击“读取选中 LaTeX”，会按 tags -> alt text -> localStorage 的顺序回填公式输入框。

如果修改代码后 PowerPoint 里仍显示旧界面，先完全退出 PowerPoint，再清理缓存：

```bash
rm -rf ~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Wef
```

受 Office.js 能力限制，原生公式依赖本地 helper 和 macOS 自动化权限；导出原图、复制大图、裁剪复制等 Windows VSTO 专属功能暂未实现。
