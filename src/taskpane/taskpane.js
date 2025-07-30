/*
 * GPTBots Copilot for Outlook
 * 集成GPTBots API的智能邮件助手
 */

/* global document, Office, API_CONFIG, getCreateConversationUrl, getChatUrl, buildCreateConversationData, buildChatRequestData, parseCreateConversationResponse, parseChatResponse */

// 全局变量
let currentConversationId = null;
let currentEmailContent = null;
let currentApiResponse = null;

// 添加全局错误处理
window.addEventListener('error', function(e) {
  console.error('🚨 全局JavaScript错误:', e.error);
  document.getElementById("sideload-msg").innerHTML = `
    <h2>JavaScript错误:</h2>
    <p>${e.error.message}</p>
    <p>文件: ${e.filename}</p>
    <p>行号: ${e.lineno}</p>
  `;
});

// 检查Office是否可用
console.log('🔍 检查Office对象:', typeof Office !== 'undefined' ? '✅ 可用' : '❌ 不可用');

// 如果Office不可用，直接显示错误
if (typeof Office === 'undefined') {
  console.error('❌ Office.js 未加载');
  document.addEventListener('DOMContentLoaded', function() {
    document.getElementById("sideload-msg").innerHTML = `
      <h2>Office.js 未加载</h2>
      <p>请检查网络连接和Office环境</p>
      <button onclick="location.reload()">重新加载</button>
    `;
  });
} else {
  console.log('✅ Office.js 已加载，版本:', Office.context ? Office.context.requirements : '未知');

  // Office初始化
  Office.onReady((info) => {
    console.log('🚀 GPTBots Copilot 开始初始化...', info);
    console.log('📊 Office信息:', {
      host: info.host,
      platform: info.platform,
      context: Office.context
    });
    
    if (info.host === Office.HostType.Outlook) {
      console.log('✅ Outlook 环境检测成功');
      
      try {
        const sideloadMsg = document.getElementById("sideload-msg");
        const appBody = document.getElementById("app-body");
        const runButton = document.getElementById("run");
        
        console.log('🔍 DOM元素检查:', {
          sideloadMsg: sideloadMsg ? '✅' : '❌',
          appBody: appBody ? '✅' : '❌', 
          runButton: runButton ? '✅' : '❌'
        });
        
        if (sideloadMsg) sideloadMsg.style.display = "none";
        if (appBody) appBody.style.display = "flex";
        if (runButton) runButton.onclick = run;
        
        console.log('✅ UI 元素设置完成');
        
        // 初始化界面
        initializeUI();
      } catch (error) {
        console.error('❌ UI 初始化失败:', error);
        showError('界面初始化失败: ' + error.message);
      }
    } else {
      console.warn('⚠️ 非Outlook环境:', info.host);
      showError(`不支持的Office应用: ${info.host || '未知'}`);
    }
  }).catch(error => {
    console.error('❌ Office.onReady 失败:', error);
    document.getElementById("sideload-msg").innerHTML = `
      <h2>Office初始化失败</h2>
      <p>${error.message}</p>
      <button onclick="location.reload()">重新加载</button>
    `;
  });
}

/**
 * 初始化用户界面
 */
function initializeUI() {
  console.log('🎨 GPTBots Copilot UI 初始化完成');
  
  // 检查API配置
  if (typeof API_CONFIG === 'undefined') {
    console.error('❌ API_CONFIG 未加载');
    showError('API配置未加载，请刷新页面重试');
    return;
  }
  
  console.log('✅ API配置检查通过:', API_CONFIG.baseUrl);
  
  // 添加调试信息到页面
  addDebugInfo();
}

/**
 * 添加调试信息
 */
function addDebugInfo() {
  const debugInfo = document.createElement('div');
  debugInfo.id = 'debug-info';
  debugInfo.style.cssText = 'position: fixed; bottom: 10px; right: 10px; background: #f0f0f0; padding: 10px; font-size: 12px; border-radius: 5px; max-width: 200px; z-index: 1000;';
  debugInfo.innerHTML = `
    <strong>调试信息:</strong><br>
    Host: ${Office.context.host}<br>
    API: ${API_CONFIG ? '✅' : '❌'}<br>
    <button onclick="toggleDebugInfo()" style="font-size: 10px; margin-top: 5px;">切换显示</button>
  `;
  document.body.appendChild(debugInfo);
}

/**
 * 切换调试信息显示
 */
function toggleDebugInfo() {
  const debugInfo = document.getElementById('debug-info');
  if (debugInfo) {
    debugInfo.style.display = debugInfo.style.display === 'none' ? 'block' : 'none';
  }
}

/**
 * 主要运行函数 - 开始使用按钮点击事件
 */
async function run() {
  try {
    showLoading('正在读取邮件内容...');
    
    // 1. 读取邮件内容
    const emailContent = await readEmailContent();
    if (!emailContent) {
      showError('无法读取邮件内容');
      return;
    }
    
    currentEmailContent = emailContent;
    showLoading('正在分析邮件内容...');
    
    // 2. 发送到GPTBots API
    const response = await sendToGPTBotsAPI(emailContent);
    if (!response.success) {
      showError('API调用失败: ' + response.error);
      return;
    }
    
    currentApiResponse = response.message;
    
    // 3. 显示结果预览界面
    showResultPreview(emailContent, response.message);
    
  } catch (error) {
    console.error('运行过程中发生错误:', error);
    showError('处理过程中发生错误: ' + error.message);
  }
}

/**
 * 读取邮件内容
 */
async function readEmailContent() {
  return new Promise((resolve, reject) => {
    try {
  const item = Office.context.mailbox.item;
      
      if (!item) {
        reject(new Error('无法获取邮件项目'));
        return;
      }
      
      // 获取邮件基本信息
      const emailInfo = {
        subject: item.subject || '无主题',
        from: item.from ? item.from.displayName + ' <' + item.from.emailAddress + '>' : '未知发件人',
        to: item.to ? item.to.map(recipient => recipient.displayName + ' <' + recipient.emailAddress + '>').join(', ') : '未知收件人',
        dateTimeCreated: item.dateTimeCreated ? item.dateTimeCreated.toLocaleString() : '未知时间'
      };
      
      // 获取邮件正文
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const emailContent = {
            ...emailInfo,
            body: result.value || '邮件正文为空'
          };
          
          console.log('成功读取邮件内容:', emailContent);
          resolve(emailContent);
        } else {
          console.error('读取邮件正文失败:', result.error);
          // 即使正文读取失败，也返回基本信息
          resolve({
            ...emailInfo,
            body: '无法读取邮件正文'
          });
        }
      });
      
    } catch (error) {
      console.error('读取邮件内容时发生错误:', error);
      reject(error);
    }
  });
}

/**
 * 发送邮件内容到GPTBots API
 */
async function sendToGPTBotsAPI(emailContent) {
  try {
    // 1. 首先创建对话
    console.log('创建对话...');
    const conversationResponse = await createConversation();
    if (!conversationResponse.success) {
      return conversationResponse;
    }
    
    currentConversationId = conversationResponse.conversationId;
    console.log('对话创建成功，ID:', currentConversationId);
    
    // 2. 构建消息内容
    const message = `请分析以下邮件内容并提供智能建议：

邮件主题: ${emailContent.subject}
发件人: ${emailContent.from}
收件人: ${emailContent.to}
发送时间: ${emailContent.dateTimeCreated}

邮件正文:
${emailContent.body}

请提供：
1. 邮件内容摘要
2. 建议的回复要点
3. 需要注意的关键信息`;

    // 3. 发送消息
    console.log('发送消息到GPTBots...');
    const chatResponse = await sendChatMessage(currentConversationId, message);
    
    return chatResponse;
    
  } catch (error) {
    console.error('GPTBots API调用失败:', error);
    return {
      success: false,
      error: error.message || '未知错误'
    };
  }
}

/**
 * 创建对话
 */
async function createConversation() {
  try {
    const url = getCreateConversationUrl();
    const data = buildCreateConversationData();
    
    console.log('🔗 创建对话请求:', {
      url: url,
      method: 'POST',
      headers: API_CONFIG.headers,
      data: data
    });
    
    const response = await fetch(url, {
      method: 'POST',
      headers: API_CONFIG.headers,
      body: JSON.stringify(data)
    });
    
    console.log('📡 HTTP响应状态:', response.status, response.statusText);
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error('❌ HTTP错误响应内容:', errorText);
      throw new Error(`HTTP错误: ${response.status} - ${response.statusText}\n响应内容: ${errorText}`);
    }
    
    const result = await response.json();
    console.log('✅ 创建对话响应:', result);
    
    const parsed = parseCreateConversationResponse(result);
    console.log('🔍 解析后的对话结果:', parsed);
    
    return parsed;
    
  } catch (error) {
    console.error('❌ 创建对话失败:', error);
    return {
      success: false,
      error: error.message || '创建对话失败'
    };
  }
}

/**
 * 发送聊天消息
 */
async function sendChatMessage(conversationId, message) {
  try {
    const url = getChatUrl();
    const messages = [{
      role: 'user',
      content: message
    }];
    const data = buildChatRequestData(conversationId, messages);
    
    console.log('💬 发送消息请求:', {
      url: url,
      conversationId: conversationId,
      messageLength: message.length,
      data: data
    });
    
    const response = await fetch(url, {
      method: 'POST',
      headers: API_CONFIG.headers,
      body: JSON.stringify(data)
    });
    
    console.log('📡 消息HTTP响应状态:', response.status, response.statusText);
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error('❌ 消息HTTP错误响应内容:', errorText);
      throw new Error(`HTTP错误: ${response.status} - ${response.statusText}\n响应内容: ${errorText}`);
    }
    
    const result = await response.json();
    console.log('✅ 消息API响应:', result);
    
    const parsed = parseChatResponse(result);
    console.log('🔍 解析后的消息结果:', parsed);
    
    return parsed;
    
  } catch (error) {
    console.error('❌ 发送消息失败:', error);
    return {
      success: false,
      error: error.message || '发送消息失败'
    };
  }
}

/**
 * 显示加载状态
 */
function showLoading(message) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main">
      <div class="ms-Spinner ms-Spinner--large">
        <div class="ms-Spinner-circle"></div>
      </div>
      <h2 class="ms-font-xl" style="margin-top: 20px;">${message}</h2>
    </div>
  `;
}

/**
 * 显示错误信息
 */
function showError(message) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main">
      <div class="ms-MessageBar ms-MessageBar--error">
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-icon">
            <i class="ms-Icon ms-Icon--ErrorBadge"></i>
          </div>
          <div class="ms-MessageBar-text">
            <span class="ms-fontWeight-semibold">错误：</span> ${message}
          </div>
        </div>
      </div>
      <div role="button" class="ms-Button ms-Button--primary" onclick="location.reload()" style="margin-top: 20px;">
        <span class="ms-Button-label">重新开始</span>
      </div>
    </div>
  `;
}

/**
 * 显示结果预览界面
 */
function showResultPreview(emailContent, apiResponse) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main" style="padding: 20px;">
      <h2 class="ms-font-xl">AI分析结果</h2>
      
      <!-- 邮件摘要 -->
      <div class="email-summary" style="background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <h3 class="ms-font-l">邮件信息</h3>
        <p><strong>主题:</strong> ${emailContent.subject}</p>
        <p><strong>发件人:</strong> ${emailContent.from}</p>
        <p><strong>时间:</strong> ${emailContent.dateTimeCreated}</p>
      </div>
      
      <!-- AI回复内容 -->
      <div class="ai-response" style="background: #fff; border: 1px solid #e1e5e9; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <h3 class="ms-font-l">AI建议</h3>
        <div style="white-space: pre-wrap; line-height: 1.6;">${apiResponse}</div>
      </div>
      
      <!-- 交互按钮 -->
      <div class="action-buttons" style="margin-top: 20px;">
        <div role="button" id="reply-button" class="ms-Button ms-Button--primary" style="margin: 5px;">
          <span class="ms-Button-label">生成回复</span>
        </div>
        <div role="button" id="forward-button" class="ms-Button ms-Button--primary" style="margin: 5px;">
          <span class="ms-Button-label">转发建议</span>
        </div>
        <div role="button" id="save-draft-button" class="ms-Button ms-Button--primary" style="margin: 5px;">
          <span class="ms-Button-label">保存草稿</span>
        </div>
        <div role="button" id="back-button" class="ms-Button" style="margin: 5px;">
          <span class="ms-Button-label">返回主页</span>
        </div>
      </div>
    </div>
  `;
  
  // 绑定按钮事件
  document.getElementById("reply-button").onclick = generateReply;
  document.getElementById("forward-button").onclick = generateForward;
  document.getElementById("save-draft-button").onclick = saveDraft;
  document.getElementById("back-button").onclick = () => location.reload();
}

/**
 * 生成回复
 */
async function generateReply() {
  try {
    showLoading('正在生成回复建议...');
    
    const replyMessage = `基于之前分析的邮件，请生成一个专业、礼貌的回复邮件内容。邮件主题是："${currentEmailContent.subject}"，发件人是："${currentEmailContent.from}"。请提供完整的回复内容，包括适当的称呼和结尾。`;
    
    const response = await sendChatMessage(currentConversationId, replyMessage);
    
    if (response.success) {
      showReplyResult(response.message);
    } else {
      showError('生成回复失败: ' + response.error);
    }
    
  } catch (error) {
    console.error('生成回复时发生错误:', error);
    showError('生成回复时发生错误: ' + error.message);
  }
}

/**
 * 生成转发建议
 */
async function generateForward() {
  try {
    showLoading('正在生成转发建议...');
    
    const forwardMessage = `基于之前分析的邮件，请提供转发建议，包括：1. 适合转发给谁 2. 转发时需要添加的说明文字 3. 需要注意的事项。`;
    
    const response = await sendChatMessage(currentConversationId, forwardMessage);
    
    if (response.success) {
      showForwardResult(response.message);
    } else {
      showError('生成转发建议失败: ' + response.error);
    }
    
  } catch (error) {
    console.error('生成转发建议时发生错误:', error);
    showError('生成转发建议时发生错误: ' + error.message);
  }
}

/**
 * 保存草稿
 */
async function saveDraft() {
  try {
    showLoading('正在保存AI分析结果到草稿...');
    
    // 构建草稿内容
    const draftContent = `GPTBots AI分析结果

原邮件信息：
主题: ${currentEmailContent.subject}
发件人: ${currentEmailContent.from}
时间: ${currentEmailContent.dateTimeCreated}

AI建议：
${currentApiResponse}

---
此内容由GPTBots Copilot生成
`;

    // 创建草稿
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [],
      subject: `AI分析: ${currentEmailContent.subject}`,
      htmlBody: draftContent.replace(/\n/g, '<br>')
    });
    
    showSuccess('草稿已创建，请查看Outlook草稿箱');
    
  } catch (error) {
    console.error('保存草稿时发生错误:', error);
    showError('保存草稿时发生错误: ' + error.message);
  }
}

/**
 * 显示回复结果
 */
function showReplyResult(replyContent) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main" style="padding: 20px;">
      <h2 class="ms-font-xl">生成的回复内容</h2>
      
      <div class="reply-content" style="background: #fff; border: 1px solid #e1e5e9; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <div style="white-space: pre-wrap; line-height: 1.6;">${replyContent}</div>
      </div>
      
      <div class="action-buttons" style="margin-top: 20px;">
        <div role="button" id="use-reply-button" class="ms-Button ms-Button--primary" style="margin: 5px;">
          <span class="ms-Button-label">使用此回复</span>
        </div>
        <div role="button" id="back-to-result-button" class="ms-Button" style="margin: 5px;">
          <span class="ms-Button-label">返回分析结果</span>
        </div>
      </div>
    </div>
  `;
  
  document.getElementById("use-reply-button").onclick = () => useReplyContent(replyContent);
  document.getElementById("back-to-result-button").onclick = () => showResultPreview(currentEmailContent, currentApiResponse);
}

/**
 * 显示转发结果
 */
function showForwardResult(forwardContent) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main" style="padding: 20px;">
      <h2 class="ms-font-xl">转发建议</h2>
      
      <div class="forward-content" style="background: #fff; border: 1px solid #e1e5e9; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <div style="white-space: pre-wrap; line-height: 1.6;">${forwardContent}</div>
      </div>
      
      <div class="action-buttons" style="margin-top: 20px;">
        <div role="button" id="back-to-result-button" class="ms-Button" style="margin: 5px;">
          <span class="ms-Button-label">返回分析结果</span>
        </div>
      </div>
    </div>
  `;
  
  document.getElementById("back-to-result-button").onclick = () => showResultPreview(currentEmailContent, currentApiResponse);
}

/**
 * 使用回复内容
 */
function useReplyContent(replyContent) {
  try {
    // 创建回复邮件
    Office.context.mailbox.item.displayReplyForm(replyContent);
    showSuccess('回复窗口已打开，内容已填入');
  } catch (error) {
    console.error('创建回复时发生错误:', error);
    showError('创建回复时发生错误: ' + error.message);
  }
}

/**
 * 显示成功信息
 */
function showSuccess(message) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main">
      <div class="ms-MessageBar ms-MessageBar--success">
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-icon">
            <i class="ms-Icon ms-Icon--Completed"></i>
          </div>
          <div class="ms-MessageBar-text">
            <span class="ms-fontWeight-semibold">成功：</span> ${message}
          </div>
        </div>
      </div>
      <div role="button" class="ms-Button ms-Button--primary" onclick="location.reload()" style="margin-top: 20px;">
        <span class="ms-Button-label">返回主页</span>
      </div>
    </div>
  `;
}

// 导出函数以供外部使用
window.run = run;
window.toggleDebugInfo = toggleDebugInfo;

// 添加全局调试函数
window.debugGPTBots = {
  testAPI: async function() {
    console.log('🧪 开始API测试...');
    try {
      const conversation = await createConversation();
      console.log('测试结果 - 创建对话:', conversation);
      
      if (conversation.success) {
        const chatResult = await sendChatMessage(conversation.conversationId, '测试消息');
        console.log('测试结果 - 发送消息:', chatResult);
      }
    } catch (error) {
      console.error('API测试失败:', error);
    }
  },
  
  showConfig: function() {
    console.log('📋 当前API配置:', API_CONFIG);
  },
  
  testEmail: async function() {
    console.log('📧 开始邮件读取测试...');
    try {
      const emailContent = await readEmailContent();
      console.log('邮件读取测试结果:', emailContent);
    } catch (error) {
      console.error('邮件读取测试失败:', error);
    }
  }
};

console.log('🔧 调试工具已加载! 使用方法:');
console.log('  debugGPTBots.testAPI() - 测试API连接');
console.log('  debugGPTBots.showConfig() - 显示配置');
console.log('  debugGPTBots.testEmail() - 测试邮件读取');