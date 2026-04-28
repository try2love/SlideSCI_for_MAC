import { writeFile } from "node:fs/promises";

const versionArg = process.argv[2];
const addinBaseUrl = (process.argv[3] || "https://try2love.github.io/SlideSCI_for_MAC").replace(/\/+$/, "");
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
2. 确保 Microsoft PowerPoint 已完全退出
3. 双击运行 \`install-slidesci-mac.command\`
4. 按系统提示允许脚本运行，并完成安装
5. 重新打开 PowerPoint
6. 在 **“视图”** 选项卡右侧找到 **SlideSCI**

## 文件说明

- \`${artifactName}\`：给最终用户使用的安装包
- \`${checksumName}\`：安装包校验文件，可选
- \`Source code\`：GitHub 自动生成的源码包，普通用户不用下载

## 当前前提

- 当前版本仍要求本机已安装 Node.js，因为本地公式 helper 还通过 Node 运行
- 加载项前端页面地址：${addinBaseUrl}

## 卸载

解压安装包后，双击运行 \`uninstall-slidesci-mac.command\`，然后完全退出并重启 PowerPoint。
`;

await writeFile(outputPath, body, "utf8");
