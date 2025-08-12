// Vercel Serverless Function - 创建对话代理
export default async function handler(req, res) {
  // 设置CORS头部 - 修复版本
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  res.setHeader('Access-Control-Max-Age', '86400'); // 24小时

  // 处理预检请求
  if (req.method === 'OPTIONS') {
    console.log('📋 处理CORS预检请求');
    res.status(200).end();
    return;
  }

  // 只允许POST请求
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    console.log('📞 代理创建对话请求...');
    
    // 转发请求到GPTBots API (新加坡端点)
    const response = await fetch('https://api-sg.gptbots.ai/v1/conversation', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': req.headers.authorization || ''
      },
      body: JSON.stringify(req.body)
    });

    const data = await response.json();
    
    if (!response.ok) {
      console.error('❌ GPTBots API错误:', response.status, data);
      return res.status(response.status).json(data);
    }

    console.log('✅ 对话创建成功:', data.data?.conversation_id);
    res.status(200).json(data);
    
  } catch (error) {
    console.error('❌ 代理服务错误:', error);
    res.status(500).json({ 
      error: 'Proxy server error', 
      message: error.message 
    });
  }
}
