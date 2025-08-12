# 🔍 调试检查清单

## 请告诉我以下信息：

### 1. **具体错误信息**
- 在Outlook插件中看到什么错误提示？
- 浏览器控制台有什么错误信息？

### 2. **测试步骤**
- 您点击了哪个功能？（深度翻译/生成摘要/生成回复）
- 在哪一步失败了？

### 3. **快速测试**
请在浏览器控制台运行以下测试，告诉我结果：

```javascript
// 测试1: 检查API配置
console.log('API配置:', API_CONFIG.baseUrl);

// 测试2: 测试Vercel代理
fetch('https://gpt-bots-mail-gskhi81x1-jackylees-projects-b81f52c7.vercel.app/api/conversation', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer app-6GQY5ONwN73Spp7Li9Bz8o37'
  },
  body: JSON.stringify({user_id: 'test-user'})
}).then(r => r.json()).then(console.log).catch(console.error);
```

### 4. **环境检查**
- 您是在哪里测试的？（Outlook Web/Outlook Desktop）
- 代码是否已经推送到GitHub Pages并更新？

## 常见问题排查：

### ❓ 如果看到 "Failed to fetch"
- 可能是Vercel代理有问题
- 需要检查API密钥是否正确

### ❓ 如果看到 "403 Forbidden"
- API密钥权限问题
- 需要验证API密钥是否有效

### ❓ 如果功能没有响应
- 可能是GitHub Pages还没有更新
- 需要等待部署完成

### ❓ 如果界面显示错误
- 检查浏览器控制台的具体错误信息
- 可能是代码加载问题
