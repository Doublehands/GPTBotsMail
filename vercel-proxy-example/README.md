# GPTBots API代理服务

这是一个使用Vercel Serverless Functions的GPTBots API代理服务，用于解决CORS跨域问题。

## 🚀 快速部署

### 1. 创建Vercel项目

1. 访问 [vercel.com](https://vercel.com)
2. 使用GitHub账号登录
3. 点击 "New Project"
4. 导入这个代理服务代码

### 2. 部署步骤

```bash
# 1. 克隆或创建项目文件夹
mkdir gptbots-proxy
cd gptbots-proxy

# 2. 复制所有文件到项目文件夹
# - api/conversation.js
# - api/message.js  
# - package.json
# - vercel.json

# 3. 安装Vercel CLI (可选)
npm i -g vercel

# 4. 部署到Vercel
vercel --prod
```

### 3. 获取部署URL

部署完成后，您会获得类似以下的URL：
```
https://your-project-name.vercel.app
```

### 4. 更新Outlook插件配置

在您的 `api-config.js` 中更新 `baseUrl`：

```javascript
const API_CONFIG = {
    baseUrl: 'https://your-project-name.vercel.app/api',
    // ... 其他配置保持不变
};
```

## 📋 API端点

- **创建对话**: `POST /api/conversation`
- **发送消息**: `POST /api/message`

## 🔧 工作原理

```
Outlook插件 → Vercel代理 → GPTBots API → Vercel代理 → Outlook插件
```

1. Outlook插件发送请求到Vercel代理
2. Vercel代理转发请求到GPTBots API
3. GPTBots API返回响应给Vercel代理
4. Vercel代理返回响应给Outlook插件

## ✅ 优势

- ✅ **完全免费** - Vercel免费额度足够demo使用
- ✅ **零CORS问题** - 完美解决跨域限制
- ✅ **快速部署** - 5分钟内完成部署
- ✅ **自动HTTPS** - 安全的SSL连接
- ✅ **全球CDN** - 快速响应速度

## 🛠️ 故障排除

如果遇到问题，检查：

1. **API密钥是否正确** - 确保在Outlook插件中配置了正确的API密钥
2. **URL是否正确** - 确保baseUrl指向您的Vercel部署URL
3. **网络连接** - 确保可以访问Vercel服务

## 📞 测试API

部署完成后，可以使用curl测试：

```bash
# 测试创建对话
curl -X POST https://your-project-name.vercel.app/api/conversation \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer your-api-key" \
  -d '{"user_id": "test-user"}'

# 测试发送消息  
curl -X POST https://your-project-name.vercel.app/api/message \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer your-api-key" \
  -d '{"conversation_id": "conv-id", "query": "Hello", "response_mode": "blocking", "user": "test-user"}'
```
