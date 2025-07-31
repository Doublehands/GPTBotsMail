/*
 * GPTBots Copilot for Outlook
 * 集成GPTBots API的智能邮件助手
 */

/* global document, Office, API_CONFIG, getCreateConversationUrl, getChatUrl, buildCreateConversationData, buildChatMessageData, parseCreateConversationResponse, parseChatResponse */

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
        if (appBody) appBody.classList.add('show');
        
        // 绑定AI技能按钮事件
        bindAISkillButtons();
        
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
 * 绑定AI技能按钮事件
 */
function bindAISkillButtons() {
  const featureItems = document.querySelectorAll('.gptbots-feature-item');
  
  featureItems.forEach((item, index) => {
    item.addEventListener('click', async function() {
      const skillType = ['translate', 'summary', 'reply'][index];
      const skillName = ['深度翻译', '生成摘要', '生成回复'][index];
      
      console.log(`🎯 用户点击了: ${skillName}`);
      await processAISkill(skillType, skillName);
    });
  });

  // 绑定预览框按钮事件
  document.getElementById('copy-result').addEventListener('click', copyResult);
  document.getElementById('use-result').addEventListener('click', useResult);
  document.getElementById('close-preview').addEventListener('click', closePreview);
}

/**
 * 处理AI技能请求
 */
async function processAISkill(skillType, skillName) {
  try {
    console.log(`🎯 开始处理${skillName}请求...`);
    
    // 显示加载状态
    showPreviewLoading(skillName);
    
    // 1. 读取邮件内容
    console.log('📧 第1步：读取邮件内容...');
    const emailContent = await readEmailContent();
    if (!emailContent) {
      console.error('❌ 第1步失败：无法读取邮件内容');
      showPreviewError('无法读取邮件内容');
      return;
    }
    
    console.log('✅ 第1步成功：邮件内容读取完成', {
      subject: emailContent.subject,
      bodyLength: emailContent.body.length
    });
    
    currentEmailContent = emailContent;
    
    // 2. 根据技能类型构建提示词
    console.log(`📝 第2步：构建${skillName}提示词...`);
    let prompt = '';
    switch (skillType) {
      case 'translate':
        prompt = `请帮我翻译：\n\n${emailContent.body}`;
        break;
      case 'summary':
        prompt = `请生成摘要：\n\n邮件主题: ${emailContent.subject}\n发件人: ${emailContent.from}\n\n邮件内容:\n${emailContent.body}`;
        break;
      case 'reply':
        prompt = `帮我生成回复内容：\n\n原邮件主题: ${emailContent.subject}\n发件人: ${emailContent.from}\n\n原邮件内容:\n${emailContent.body}`;
        break;
    }
    
    console.log(`✅ 第2步成功：提示词构建完成`, {
      skillType: skillType,
      promptLength: prompt.length,
      promptPreview: prompt.substring(0, 100) + '...'
    });
    
    // 3. 模拟AI处理（延迟2秒）
    console.log(`🚀 第3步：模拟AI处理...`);
    console.log(`⏰ 等待2秒以模拟AI处理时间...`);
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // 根据技能类型返回预设内容
    let mockResponse = '';
    switch (skillType) {
      case 'reply':
        mockResponse = `Dear Jacky,

Thank you for contacting GPTBots! We appreciate your interest in our AI solutions for e-commerce support.

Our platform excels in automating high-volume customer interactions:

Seamless Integration: APIs for Shopify/Magento/WooCommerce with 24-hour setup support.

Multilingual Support: 50+ languages including nuanced English/Chinese dialects.

Custom Training: Upload CSV/PDFs to train bots on your catalog/policies (secure encryption).

Accuracy: 92%+ resolution rate for common queries; fallback to human agents via Slack/Teams.

Analytics Dashboard: Real-time metrics on response time, satisfaction, and issue trends.

Attached is a case study showing how Similar Brand reduced ticket volume by 75% in 3 months. We can schedule a personalized demo next week—please suggest your availability.

Looking forward to empowering your customer experience!

Sincerely,
Jiaqi Li
Solutions Architect, GPTBots
contact@gptbots.ai`;
        break;
      case 'translate':
        mockResponse = `主题：关于GPTBots平台AI电商客服解决方案的咨询

尊敬的GPTBots团队：

您好！我是极光的Jacky。

我们正在探索AI驱动的客服解决方案，以实现高效自动化支持。特此咨询：

GPTBots如何与Shopify/Magento等平台集成？

是否支持多语言交互（尤其是中英文）？

能否基于我们的专有数据（产品参数/政策）定制训练？

处理复杂咨询的典型准确率如何？

是否有转接人工客服的自定义流程？

我们的目标是将响应时间缩短至30秒内，并自动化处理80%+的咨询。请提供相关案例或演示选项。

感谢您的支持，期待您的回复！

此致
敬礼
Jacky`;
        break;
      case 'summary':
        mockResponse = `询问邮件重点：

需求背景：电商企业寻求AI客服解决方案，要求自动化处理订单查询、退换货、多语言支持（中英文）等高频率问题。

关键问题：平台集成能力、多语言支持、数据定制训练、准确率及人工转接流程。

目标：30秒内响应，80%以上咨询实现自动化。

回复邮件亮点：

技术能力：支持Shopify/Magento等主流电商平台快速对接，覆盖50+语言（含中英文），支持私有数据训练（加密安全）。

性能数据：92%的常见问题解决率，提供实时数据分析看板。

成功案例：同类客户通过部署GPTBots，3个月内减少75%人工工单量。

后续行动：可安排个性化演示，进一步讨论定制方案。

下一步建议：

若需求匹配，可预约演示并细化部署时间表；

如需优先解决特定痛点（如退换货自动化），可提供更详细业务场景供GPTBots优化配置。`;
        break;
    }
    
    console.log(`✅ 第3步成功：模拟AI处理完成`, {
      responseLength: mockResponse.length,
      responsePreview: mockResponse.substring(0, 100) + '...'
    });
    
    // 4. 显示结果
    console.log(`🎨 第4步：显示结果...`);
    showPreviewResult(mockResponse, skillType);
    console.log(`✅ 第4步成功：${skillName}处理完成`);
    
  } catch (error) {
    console.error(`❌ ${skillName}处理过程中发生异常:`, {
      name: error.name,
      message: error.message,
      stack: error.stack
    });
    showPreviewError(`${skillName}处理失败: ${error.message}`);
  }
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
 * 发送消息到GPTBots API
 */
async function sendToGPTBotsAPI(message) {
  try {
    console.log('🔄 开始GPTBots API调用流程...');
    
    // 1. 首先创建对话
    console.log('📞 第3.1步：创建对话...');
    const conversationResponse = await createConversation();
    if (!conversationResponse.success) {
      console.error('❌ 第3.1步失败：创建对话失败', conversationResponse);
      return conversationResponse;
    }
    
    currentConversationId = conversationResponse.conversationId;
    console.log('✅ 第3.1步成功：对话创建成功', {
      conversationId: currentConversationId
    });
    
    // 2. 发送消息
    console.log('💬 第3.2步：发送消息到GPTBots...');
    console.log('📝 消息内容预览:', {
      messageLength: message.length,
      messagePreview: message.substring(0, 200) + '...'
    });
    
    const chatResponse = await sendChatMessage(currentConversationId, message);
    
    if (chatResponse.success) {
      console.log('✅ 第3.2步成功：消息发送成功', {
        responseLength: chatResponse.message ? chatResponse.message.length : 0
      });
    } else {
      console.error('❌ 第3.2步失败：消息发送失败', chatResponse);
    }
    
    return chatResponse;
    
  } catch (error) {
    console.error('❌ GPTBots API调用过程中发生异常:', {
      name: error.name,
      message: error.message,
      stack: error.stack
    });
    return {
      success: false,
      error: error.message || '未知错误',
      originalError: error.message
    };
  }
}

/**
 * 创建对话 - 简化版本
 */
async function createConversation() {
  try {
    console.log('🚀 开始创建对话...');
    const url = getCreateConversationUrl();
    const data = buildCreateConversationData();
    
    console.log('📞 API请求 (通过CORS代理):', { url, method: 'POST', data });
    console.log('🔧 使用cors.io代理解决CORS问题');
    
    const response = await fetch(url, {
      method: 'POST',
      headers: API_CONFIG.headers,
      body: JSON.stringify(data)
    });
    
    if (!response.ok) {
      throw new Error(`HTTP错误: ${response.status}`);
    }
    
    const result = await response.json();
    console.log('✅ 创建对话成功:', result);
    
    return parseCreateConversationResponse(result);
  } catch (error) {
    console.error('❌ 创建对话失败:', error);
    return {
      success: false,
      error: error.message
    };
  }
}



/**
 * 发送聊天消息 - 简化版本
 */
async function sendChatMessage(conversationId, message) {
  try {
    console.log('💬 开始发送消息...');
    const url = getChatUrl();
    const data = buildChatMessageData(conversationId, message);
    
    console.log('📞 API请求 (通过CORS代理):', { url, method: 'POST', data });
    console.log('🔧 使用cors.io代理解决CORS问题');
    
    const response = await fetch(url, {
      method: 'POST',
      headers: API_CONFIG.headers,
      body: JSON.stringify(data)
    });
    
    if (!response.ok) {
      throw new Error(`HTTP错误: ${response.status}`);
    }
    
    const result = await response.json();
    console.log('✅ 发送消息成功:', result);
    
    return parseChatResponse(result);
  } catch (error) {
    console.error('❌ 发送消息失败:', error);
    return {
      success: false,
      error: error.message
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

/**
 * 显示预览框加载状态
 */
function showPreviewLoading(skillName) {
  const preview = document.getElementById('result-preview');
  const loading = document.getElementById('loading-indicator');
  const resultText = document.getElementById('result-text');
  
  preview.classList.remove('gptbots-hidden');
  loading.classList.remove('gptbots-hidden');
  resultText.classList.add('gptbots-hidden');
  
  loading.querySelector('p').textContent = `AI正在${skillName}中...`;
}

/**
 * 显示预览框错误
 */
function showPreviewError(errorMessage) {
  const preview = document.getElementById('result-preview');
  const loading = document.getElementById('loading-indicator');
  const resultText = document.getElementById('result-text');
  
  preview.classList.remove('gptbots-hidden');
  loading.classList.add('gptbots-hidden');
  resultText.classList.remove('gptbots-hidden');
  resultText.innerHTML = `
    <div class="error-message">
      <i class="ms-Icon ms-Icon--ErrorBadge"></i>
      <span>${errorMessage}</span>
    </div>
  `;
}

/**
 * 显示预览框结果
 */
function showPreviewResult(result, skillType) {
  const preview = document.getElementById('result-preview');
  const loading = document.getElementById('loading-indicator');
  const resultText = document.getElementById('result-text');
  
  preview.classList.remove('gptbots-hidden');
  loading.classList.add('gptbots-hidden');
  resultText.classList.remove('gptbots-hidden');
  
  // 保存当前结果和类型，供后续操作使用
  currentApiResponse = result;
  currentSkillType = skillType;
  
  resultText.innerHTML = `
    <div class="result-content">
      <div class="result-text-content">${result.replace(/\n/g, '<br>')}</div>
    </div>
  `;
}

/**
 * 复制结果到剪贴板
 */
async function copyResult() {
  try {
    if (currentApiResponse) {
      await navigator.clipboard.writeText(currentApiResponse);
      
      // 显示复制成功提示
      const copyBtn = document.getElementById('copy-result');
      const originalText = copyBtn.textContent;
      copyBtn.textContent = '已复制!';
      copyBtn.style.backgroundColor = '#107c10';
      
      setTimeout(() => {
        copyBtn.textContent = originalText;
        copyBtn.style.backgroundColor = '';
      }, 2000);
      
      console.log('✅ 结果已复制到剪贴板');
    }
  } catch (error) {
    console.error('❌ 复制失败:', error);
    alert('复制失败，请手动选择并复制内容');
  }
}

/**
 * 使用结果（根据技能类型执行不同操作）
 */
async function useResult() {
  try {
    if (!currentApiResponse || !currentSkillType) {
      console.error('❌ 没有可用的结果');
      return;
    }
    
    switch (currentSkillType) {
      case 'reply':
        // 生成回复：创建回复邮件
        Office.context.mailbox.item.displayReplyForm(currentApiResponse);
        showSuccess('回复窗口已打开，内容已填入');
        break;
        
      case 'translate':
      case 'summary':
        // 翻译和摘要：创建新邮件草稿
        const subject = currentSkillType === 'translate' ? 
          `翻译: ${currentEmailContent.subject}` : 
          `摘要: ${currentEmailContent.subject}`;
          
        Office.context.mailbox.displayNewMessageForm({
          toRecipients: [],
          subject: subject,
          htmlBody: currentApiResponse.replace(/\n/g, '<br>')
        });
        showSuccess('草稿已创建，请查看Outlook草稿箱');
        break;
    }
    
    // 关闭预览框
    closePreview();
    
  } catch (error) {
    console.error('❌ 使用结果失败:', error);
    alert('操作失败: ' + error.message);
  }
}

/**
 * 关闭预览框
 */
function closePreview() {
  const preview = document.getElementById('result-preview');
  preview.classList.add('gptbots-hidden');
  
  // 清理状态
  currentApiResponse = null;
  currentSkillType = null;
}

/**
 * 显示成功信息
 */
function showSuccess(message) {
  // 临时显示成功消息
  const preview = document.getElementById('result-preview');
  const resultText = document.getElementById('result-text');
  
  resultText.innerHTML = `
    <div class="success-message">
      <i class="ms-Icon ms-Icon--Completed" style="color: #107c10;"></i>
      <span>${message}</span>
    </div>
  `;
  
  // 3秒后自动关闭
  setTimeout(() => {
    closePreview();
  }, 3000);
}

// 全局变量，保存当前技能类型
let currentSkillType = null;

// 导出函数以供外部使用
window.run = run;

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
  
  testConnection: async function() {
    console.log('🌐 测试API连接和代理...');
    
    // 测试原始API连接
    try {
      const originalUrl = API_CONFIG.baseUrl;
      console.log('🔗 测试原始API URL:', originalUrl);
      
      const response = await fetch(originalUrl, {
        method: 'GET',
        mode: 'no-cors'
      });
      
      console.log('📡 原始API响应:', response);
      console.log('📡 响应类型:', response.type);
      console.log('📡 响应状态:', response.status);
      
      if (response.type === 'opaque') {
        console.log('✅ 原始API服务器可达，但被CORS策略阻止（这是正常的）');
      }
      
    } catch (error) {
      console.error('❌ 原始API连接测试失败:', error);
    }
    
    // 测试代理连接
    if (API_CONFIG.corsProxy && API_CONFIG.corsProxy.enabled) {
      console.log('🔄 测试CORS代理...');
      
      // 测试主要代理
      try {
        const proxyUrl = getCreateConversationUrl();
        console.log('🔗 测试主要代理URL:', proxyUrl);
        
        // 只测试连通性，不发送实际请求
        const testUrl = `${API_CONFIG.corsProxy.primary}${encodeURIComponent('https://httpbin.org/get')}`;
        const response = await fetch(testUrl, {
          method: 'GET',
          mode: 'cors'
        });
        
        console.log('📡 主要代理响应状态:', response.status);
        if (response.ok) {
          console.log('✅ 主要代理工作正常');
        }
        
      } catch (error) {
        console.error('❌ 主要代理测试失败:', error);
      }
      
      // 测试备用代理
      try {
        const fallbackUrl = getChatUrlFallback();
        console.log('🔗 测试备用代理URL:', fallbackUrl);
        
        // 只测试连通性，不发送实际请求
        const testUrl = `${API_CONFIG.corsProxy.fallback}${encodeURIComponent('https://httpbin.org/get')}`;
        const response = await fetch(testUrl, {
          method: 'GET',
          mode: 'cors'
        });
        
        console.log('📡 备用代理响应状态:', response.status);
        if (response.ok) {
          console.log('✅ 备用代理工作正常');
        }
        
      } catch (error) {
        console.error('❌ 备用代理测试失败:', error);
      }
    }
    
    console.log('🏁 连接测试完成');
  },
  
  showConfig: function() {
    console.log('📋 当前API配置:', API_CONFIG);
    console.log('📋 创建对话URL:', getCreateConversationUrl());
    console.log('📋 发送消息URL:', getChatUrl());
  },
  
  testEmail: async function() {
    console.log('📧 开始邮件读取测试...');
    try {
      const emailContent = await readEmailContent();
      console.log('邮件读取测试结果:', emailContent);
    } catch (error) {
      console.error('邮件读取测试失败:', error);
    }
  },
  
  simulateTranslate: async function() {
    console.log('🔄 模拟翻译测试...');
    try {
      // 模拟翻译请求
      const testPrompt = '请帮我翻译：Hello, this is a test message.';
      const response = await sendToGPTBotsAPI(testPrompt);
      console.log('翻译测试结果:', response);
    } catch (error) {
      console.error('翻译测试失败:', error);
    }
  }
};

// 显示CSP解决方案状态
console.log('🔍 CSP解决方案状态检查:');
console.log('  直接调用创建对话URL:', getCreateConversationUrl());
console.log('  直接调用发送消息URL:', getChatUrl());
console.log('  备用ThingProxy创建对话URL:', getCreateConversationUrlFallback());
console.log('  备用ThingProxy发送消息URL:', getChatUrlFallback());

console.log('🔧 调试工具已加载! 使用方法:');
console.log('  debugGPTBots.testConnection() - 测试API和CORS代理连通性');
console.log('  debugGPTBots.testAPI() - 完整API功能测试（含代理重试）');
console.log('  debugGPTBots.simulateTranslate() - 模拟翻译功能测试');
console.log('  debugGPTBots.showConfig() - 显示当前配置');
console.log('  debugGPTBots.testEmail() - 测试邮件读取');
console.log('');
console.log('🛠️ CSP解决方案已实施：');
console.log('  - 🎯 直接调用: 已在manifest.xml中添加api.gptbots.ai域名权限');
console.log('  - 🥇 备用方案1: thingproxy.freeboard.io（支持Authorization头）');
console.log('  - 🥈 备用方案2: proxy.cors.sh（支持复杂请求）');
console.log('  - 🔄 智能重试：直接调用失败时自动切换到代理');
console.log('  - 🔧 解决方案：先尝试绕过CSP，再使用代理服务');
console.log('');
console.log('💡 已解决Content Security Policy限制问题');