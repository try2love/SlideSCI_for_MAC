import { writeFile } from "node:fs/promises";

const versionArg = process.argv[2];
const addinBaseUrl = (process.argv[3] || "https://127.0.0.1:18443").replace(/\/+$/, "");
const outputPath = process.argv[4];

if (!versionArg || !outputPath) {
  console.error("Usage: node scripts/generate-release-notes.mjs <version> <addin_base_url> <output_path>");
  process.exit(1);
}

const version = versionArg.replace(/^v/, "");
const artifactName = `SlideSCI-for-Mac-v${version}.zip`;
const checksumName = `SlideSCI-for-Mac-v${version}.sha256`;

const body = `## 安装

请下载 **\`${artifactName}\`**，不要下载 GitHub 自动附带的 \`Source code\`。

1. 下载并解压 \`${artifactName}\`
2. 如果你的 Mac 还没有安装 Node.js，请先安装 Node.js LTS：<https://nodejs.org/>
3. 确保 Microsoft PowerPoint 已完全退出
4. 双击运行 \`install-slidesci-mac.command\`
5. 如果 macOS 阻止运行，请右键该文件，选择“打开”，再确认一次
6. 按系统提示允许脚本运行，并完成安装
7. 安装脚本会同时安装本地任务窗格 HTTPS 服务，并在首次安装时把本地证书加入登录钥匙串
8. 如果系统弹出权限提示，请允许终端或 Node 控制 Microsoft PowerPoint，并允许辅助功能权限
9. 重新打开 PowerPoint
10. 在 **“视图”** 选项卡右侧找到 **SlideSCI**

## 文件说明

- \`${artifactName}\`：给最终用户使用的安装包
- \`${checksumName}\`：安装包校验文件，可选
- \`Source code\`：GitHub 自动生成的源码包，普通用户不用下载

## 当前前提

- 当前版本仍要求本机已安装 Node.js，因为本地公式 helper 和本地任务窗格服务都通过 Node 运行
- 运行时不依赖 GitHub Pages；加载项页面默认由本机地址提供：${addinBaseUrl}

## 卸载

1. 完全退出 PowerPoint
2. 解压安装包后，双击运行 \`uninstall-slidesci-mac.command\`
3. 卸载完成后重新打开 PowerPoint
`;

await writeFile(outputPath, body, "utf8");
