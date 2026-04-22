这是一份包含你要求的各种 Markdown 原生格式的示例。你可以直接复制下面的内容，它展示了标准 Markdown 的各种高级语法：

---

# 一级标题：项目研究概览
## 二级标题：RAG 安全性分析
### 三级标题：技术细节与公式

这是一个包含**加粗**、*斜体*以及`行内代码`的普通段落。以下是该项目的主要组成部分。

---

### 1. 技术栈清单（无序列表）
* **核心框架**：Python / PyTorch
* **前端工具**：Vite / Office JS API
* **安全研究**：数据毒化、DoS 攻击模拟
* **开发环境**：macOS (M4 MacBook Air)

### 2. 开发进度表（任务列表）
- [x] 完成本地代理 15926 端口配置
- [x] 解决 Xcode 的 DerivedData 权限问题
- [ ] 部署 SlideSCI 插件到生产环境
- [ ] 完成向量数据库的投毒防御实验

### 3. 系统架构对比（表格）

| 模块 | 功能 | 状态 | 备注 |
| :--- | :--- | :--- | :--- |
| **Backend** | RAG 检索逻辑 | 运行中 | 部署于 88 服务器 |
| **Frontend** | Office Add-in UI | 调试中 | Vite 构建 |
| **Proxy** | SOCKS5 代理 | 待修复 | 端口 15926 |

---

### 4. 核心防御算法（数学公式）

在研究 RAG 安全时，我们定义了知识库注入的评分函数。以下是其跨行数学公式：

$$
Score(d_i, q) = \frac{\sum_{j=1}^{n} w_j \cdot cos(\vec{v}_{d_i}, \vec{v}_q)}{\sqrt{\delta + \lambda \cdot |Noise|}}
$$

其中，$\delta$ 是平滑因子，$\lambda$ 表示噪声衰减系数。

---

### 5. 快速启动脚本（代码块）

以下是用于修复权限并启动 Vite 开发环境的 Bash 命令：

```bash
# 赋予执行权限并清理 node_modules
chmod -R +x node_modules/.bin/
rm -rf node_modules package-lock.json

# 重新安装并运行
npm install
npm run dev -- --host 0.0.0.0
```

---

### 6. 重要提示（引述块）

> **注意：** 在 macOS 上侧载 Office 插件时，如果修改了 `manifest.xml` 但界面没有变化，请务必清理缓存：
> `rm -rf ~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Wef`

---

### 7. 归档步骤（有序列表）
1. 提交本地代码到 `feature` 分支。
2. 通过 `git reset --hard` 回退 `master`。
3. 发起 Pull Request 到主仓库。