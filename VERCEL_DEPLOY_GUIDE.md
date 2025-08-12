# 📚 Vercel部署详细指南

## 🚀 方法一：通过GitHub部署（推荐）

### 步骤1：准备GitHub仓库

1. **登录GitHub**
   - 访问 [github.com](https://github.com)
   - 登录您的账号

2. **创建新仓库**
   - 点击右上角的 "+" 按钮
   - 选择 "New repository"
   - 仓库名称：`gptbots-proxy`
   - 设置为 Public
   - 点击 "Create repository"

3. **上传代码文件**
   - 点击 "uploading an existing file"
   - 将以下文件拖拽上传：
     ```
     vercel-proxy-example/api/conversation.js
     vercel-proxy-example/api/message.js
     vercel-proxy-example/package.json
     vercel-proxy-example/vercel.json
     vercel-proxy-example/README.md
     ```
   - 保持文件夹结构：`api/conversation.js`, `api/message.js`
   - 点击 "Commit changes"

### 步骤2：连接Vercel

1. **访问Vercel**
   - 打开 [vercel.com](https://vercel.com)
   - 点击右上角 "Sign Up" 或 "Log In"

2. **使用GitHub登录**
   - 选择 "Continue with GitHub"
   - 授权Vercel访问您的GitHub账号

3. **导入项目**
   - 登录后，点击 "New Project"
   - 在列表中找到 `gptbots-proxy` 仓库
   - 点击 "Import"

### 步骤3：配置部署

1. **项目设置**
   - Project Name: 保持默认或改为 `gptbots-proxy`
   - Framework Preset: 选择 "Other"
   - Root Directory: 保持默认 `./`

2. **环境变量**（可选）
   - 暂时不需要设置
   - 直接点击 "Deploy"

### 步骤4：等待部署完成

1. **部署过程**
   - Vercel会自动构建和部署
   - 通常需要1-2分钟
   - 您会看到进度条和日志

2. **获取URL**
   - 部署成功后，您会看到 "Congratulations!" 页面
   - 复制显示的URL，例如：`https://gptbots-proxy-abc123.vercel.app`

## 🚀 方法二：直接上传文件

### 如果您不想使用GitHub：

1. **访问Vercel**
   - 打开 [vercel.com](https://vercel.com)
   - 注册/登录账号

2. **拖拽部署**
   - 在首页找到 "Deploy with drag and drop"
   - 将 `vercel-proxy-example` 整个文件夹拖拽到页面上
   - Vercel会自动开始部署

3. **获取URL**
   - 部署完成后复制URL

## 📝 部署后测试

### 测试API是否工作：

```bash
# 替换URL为您的实际Vercel URL
curl -X POST https://your-vercel-url.vercel.app/api/conversation \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer app-6GQY5ONwN73Spp7Li9Bz8o37" \
  -d '{"user_id": "test-user"}'
```

### 预期响应：
```json
{
  "data": {
    "conversation_id": "conv-xxxx-xxxx-xxxx"
  }
}
```

## 🔧 更新Outlook插件配置

部署成功后，将URL更新到您的 `api-config.js`：

```javascript
baseUrl: 'https://your-actual-vercel-url.vercel.app/api',
```

## ❓ 常见问题

### Q: 部署失败了怎么办？
A: 检查文件结构是否正确：
```
根目录/
├── api/
│   ├── conversation.js
│   └── message.js
├── package.json
└── vercel.json
```

### Q: 找不到GitHub仓库？
A: 确保：
- 仓库是Public的
- 已经授权Vercel访问GitHub
- 刷新页面重试

### Q: API调用失败？
A: 检查：
- URL是否正确
- API密钥是否有效
- 网络连接是否正常

## 📞 需要帮助？

如果遇到任何问题，请告诉我：
1. 在哪一步遇到了困难
2. 看到了什么错误信息
3. 截图（如果可能）

我会详细指导您解决！
