/*
 * GPTBots Copilot for Outlook
 * 集成GPTBots API的智能邮件助手
 */

/* global document, Office, API_CONFIG, getCreateConversationUrl, getChatUrl, buildCreateConversationData, buildChatRequestData, parseCreateConversationResponse, parseChatResponse */

// 全局变量
let currentConversationId = null;
let currentEmailContent = null;
let currentApiResponse = null;
let currentMode = null; // 'Read' 或 'Compose'
let previewContent = null;

// Office初始化
Office.onReady((info) => {
  console.log('🚀 GPTBots Copilot 开始初始化...', info);
  
  if (info.host === Office.HostType.Outlook) {
    console.log('✅ Outlook 环境检测成功');
    
    try {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
      
      // 检测当前模式
      detectCurrentMode();
      
      // 绑定公共按钮事件
      document.getElementById("clear-preview-btn").onclick = clearPreview;
      document.getElementById("insert-content-btn").onclick = insertContentToEmail;
      
      console.log('✅ UI 元素绑定成功');
      
      // 初始化界面
      initializeUI();
    } catch (error) {
      console.error('❌ UI 初始化失败:', error);
      showError('界面初始化失败: ' + error.message);
    }
  } else {
    console.warn('⚠️ 非Outlook环境:', info.host);
    showError(`不支持的Office应用: ${info.host}`);
  }
});

/**
 * 检测当前模式（阅读或编辑）
 */
function detectCurrentMode() {
  try {
    const item = Office.context.mailbox.item;
    
    // 通过不同方法检测模式
    if (item.addHandlerAsync && item.removeHandlerAsync) {
      // 编辑模式特有的方法
      currentMode = 'Compose';
    } else if (item.dateTimeCreated !== undefined) {
      // 阅读模式特有的属性
      currentMode = 'Read';
    } else {
      // 备用检测方法
      currentMode = item.itemType === Office.MailboxEnums.ItemType.Message ? 'Read' : 'Compose';
    }
    
    console.log('🔍 检测到当前模式:', currentMode);
    
    // 更新模式指示器
    const modeIndicator = document.getElementById('mode-indicator');
    if (modeIndicator) {
      modeIndicator.textContent = currentMode === 'Read' ? '📖 邮件阅读模式' : '✍️ 邮件编辑模式';
    }
    
  } catch (error) {
    console.error('❌ 模式检测失败:', error);
    currentMode = 'Read'; // 默认为阅读模式
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
  
  // 根据模式显示相应的按钮
  setupModeBasedUI();
  
  // 添加调试信息到页面
  addDebugInfo();
}

/**
 * 根据模式设置UI
 */
function setupModeBasedUI() {
  const readModeButtons = document.getElementById('read-mode-buttons');
  const composeModeButtons = document.getElementById('compose-mode-buttons');
  const previewSection = document.getElementById('preview-section');
  const insertBtn = document.getElementById('insert-content-btn');
  
  if (currentMode === 'Read') {
    // 阅读模式
    readModeButtons.style.display = 'flex';
    composeModeButtons.style.display = 'none';
    insertBtn.style.display = 'none';
    
    // 绑定阅读模式按钮事件
    document.getElementById("deep-translate-btn").onclick = () => handleReadModeAction('translate');
    document.getElementById("generate-summary-btn").onclick = () => handleReadModeAction('summary');
    document.getElementById("generate-reply-btn").onclick = () => handleReadModeAction('reply');
    
  } else {
    // 编辑模式
    readModeButtons.style.display = 'none';
    composeModeButtons.style.display = 'flex';
    insertBtn.style.display = 'inline-block';
    
    // 绑定编辑模式按钮事件
    document.getElementById("compose-translate-btn").onclick = () => handleComposeModeAction('translate');
    document.getElementById("content-polish-btn").onclick = () => handleComposeModeAction('polish');
    document.getElementById("compose-reply-btn").onclick = () => handleComposeModeAction('reply');
    document.getElementById("generate-draft-btn").onclick = () => handleComposeModeAction('draft');
  }
  
  // 显示预览区域
  previewSection.style.display = 'block';
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
 * 处理阅读模式的操作
 */
async function handleReadModeAction(action) {
  try {
    showPreviewLoading(`正在${getActionName(action)}...`);
    
    // 读取邮件内容
    const emailContent = await readEmailContent();
    if (!emailContent) {
      showPreviewError('无法读取邮件内容');
      return;
    }
    
    currentEmailContent = emailContent;
    
    let prompt;
    switch (action) {
      case 'translate':
        prompt = buildTranslatePrompt(emailContent);
        break;
      case 'summary':
        prompt = buildSummaryPrompt(emailContent);
        break;
      case 'reply':
        prompt = buildReplyPrompt(emailContent);
        break;
    }
    
    // 发送到GPTBots API
    const response = await sendToGPTBotsAPI(emailContent, prompt);
    if (!response.success) {
      showPreviewError(`${getActionName(action)}失败: ${response.error}`);
      return;
    }
    
    // 显示在预览框中
    showPreviewContent(response.message, action);
    
  } catch (error) {
    console.error(`${getActionName(action)}过程中发生错误:`, error);
    showPreviewError(`${getActionName(action)}过程中发生错误: ${error.message}`);
  }
}

/**
 * 处理编辑模式的操作
 */
async function handleComposeModeAction(action) {
  try {
    showPreviewLoading(`正在${getActionName(action)}...`);
    
    let content = '';
    let prompt = '';
    
    if (action === 'translate' || action === 'polish') {
      // 需要获取当前正在编辑的内容
      content = await getCurrentComposeContent();
      if (!content || content.trim() === '') {
        showPreviewError('请先在邮件中输入内容');
        return;
      }
    }
    
    switch (action) {
      case 'translate':
        prompt = buildComposeTranslatePrompt(content);
        break;
      case 'polish':
        prompt = buildPolishPrompt(content);
        break;
      case 'reply':
        // 编辑模式下的生成回复（基于主题或上下文）
        const context = await getComposeContext();
        prompt = buildComposeReplyPrompt(context);
        break;
      case 'draft':
        // 生成草稿（基于主题）
        const subject = await getCurrentSubject();
        prompt = buildDraftPrompt(subject);
        break;
    }
    
    // 发送到GPTBots API
    const response = await sendToGPTBotsAPI({ body: content }, prompt);
    if (!response.success) {
      showPreviewError(`${getActionName(action)}失败: ${response.error}`);
      return;
    }
    
    // 显示在预览框中
    showPreviewContent(response.message, action);
    
  } catch (error) {
    console.error(`${getActionName(action)}过程中发生错误:`, error);
    showPreviewError(`${getActionName(action)}过程中发生错误: ${error.message}`);
  }
}

/**
 * 清空预览内容
 */
function clearPreview() {
  const previewContentEl = document.getElementById('preview-content');
  if (previewContentEl) {
    previewContentEl.textContent = '点击上方按钮生成AI内容...';
    previewContentEl.style.color = '#666';
  }
  previewContent = null;
  
  // 隐藏插入按钮
  const insertBtn = document.getElementById('insert-content-btn');
  if (insertBtn) {
    insertBtn.style.display = 'none';
  }
}

/**
 * 插入内容到邮件编辑器
 */
async function insertContentToEmail() {
  if (!previewContent || currentMode !== 'Compose') {
    console.warn('⚠️ 无预览内容或非编辑模式');
    return;
  }
  
  try {
    showPreviewLoading('正在插入内容...');
    
    // 获取当前邮件正文
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const currentBody = result.value || '';
        
        // 如果邮件为空，直接设置内容
        if (!currentBody.trim() || currentBody.trim() === '<div></div>') {
          Office.context.mailbox.item.body.setAsync(
            previewContent,
            { coercionType: Office.CoercionType.Html },
            (setResult) => {
              if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                showPreviewSuccess('内容已插入到邮件中');
              } else {
                showPreviewError('插入内容失败: ' + setResult.error.message);
              }
            }
          );
        } else {
          // 如果有内容，在末尾添加
          const newContent = currentBody + '<br><br>' + previewContent;
          Office.context.mailbox.item.body.setAsync(
            newContent,
            { coercionType: Office.CoercionType.Html },
            (setResult) => {
              if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                showPreviewSuccess('内容已添加到邮件末尾');
              } else {
                showPreviewError('插入内容失败: ' + setResult.error.message);
              }
            }
          );
        }
      } else {
        showPreviewError('获取邮件内容失败: ' + result.error.message);
      }
    });
    
  } catch (error) {
    console.error('插入内容时发生错误:', error);
    showPreviewError('插入内容时发生错误: ' + error.message);
  }
}

/**
 * 深度翻译功能（保留原函数用于兼容）
 */
async function deepTranslate() {
  try {
    showLoading('正在读取邮件内容...');
    
    // 1. 读取邮件内容
    const emailContent = await readEmailContent();
    if (!emailContent) {
      showError('无法读取邮件内容');
      return;
    }
    
    currentEmailContent = emailContent;
    showLoading('正在进行深度翻译...');
    
    // 2. 构建翻译提示词
    const translatePrompt = `请对以下邮件进行深度翻译，保持专业性和语境准确性：

邮件主题: ${emailContent.subject}
发件人: ${emailContent.from}

邮件正文:
${emailContent.body}

请提供：
1. 完整的中文翻译
2. 关键术语解释
3. 语境背景说明（如有必要）`;

    // 3. 发送到GPTBots API
    const response = await sendToGPTBotsAPI(emailContent, translatePrompt);
    if (!response.success) {
      showError('翻译失败: ' + response.error);
      return;
    }
    
    currentApiResponse = response.message;
    showTranslationResult(emailContent, response.message);
    
  } catch (error) {
    console.error('翻译过程中发生错误:', error);
    showError('翻译过程中发生错误: ' + error.message);
  }
}

/**
 * 生成摘要功能
 */
async function generateSummary() {
  try {
    showLoading('正在读取邮件内容...');
    
    // 1. 读取邮件内容
    const emailContent = await readEmailContent();
    if (!emailContent) {
      showError('无法读取邮件内容');
      return;
    }
    
    currentEmailContent = emailContent;
    showLoading('正在生成邮件摘要...');
    
    // 2. 构建摘要提示词
    const summaryPrompt = `请为以下邮件生成详细摘要：

邮件主题: ${emailContent.subject}
发件人: ${emailContent.from}
发送时间: ${emailContent.dateTimeCreated}

邮件正文:
${emailContent.body}

请提供：
1. 邮件核心内容摘要（2-3句话）
2. 关键信息点列表
3. 重要日期和截止时间（如有）
4. 需要采取的行动（如有）
5. 优先级评估`;

    // 3. 发送到GPTBots API
    const response = await sendToGPTBotsAPI(emailContent, summaryPrompt);
    if (!response.success) {
      showError('生成摘要失败: ' + response.error);
      return;
    }
    
    currentApiResponse = response.message;
    showSummaryResult(emailContent, response.message);
    
  } catch (error) {
    console.error('生成摘要过程中发生错误:', error);
    showError('生成摘要过程中发生错误: ' + error.message);
  }
}

/**
 * 生成回复功能
 */
async function generateReply() {
  try {
    showLoading('正在读取邮件内容...');
    
    // 1. 读取邮件内容
    const emailContent = await readEmailContent();
    if (!emailContent) {
      showError('无法读取邮件内容');
      return;
    }
    
    currentEmailContent = emailContent;
    showLoading('正在生成智能回复...');
    
    // 2. 构建回复提示词
    const replyPrompt = `请为以下邮件生成专业的回复建议：

邮件主题: ${emailContent.subject}
发件人: ${emailContent.from}
发送时间: ${emailContent.dateTimeCreated}

邮件正文:
${emailContent.body}

请提供：
1. 推荐的回复内容（专业、礼貌、完整）
2. 回复要点分析
3. 语气建议（正式/非正式）
4. 需要补充的信息（如有）
5. 后续跟进建议（如需要）`;

    // 3. 发送到GPTBots API
    const response = await sendToGPTBotsAPI(emailContent, replyPrompt);
    if (!response.success) {
      showError('生成回复失败: ' + response.error);
      return;
    }
    
    currentApiResponse = response.message;
    showReplyResult(emailContent, response.message);
    
  } catch (error) {
    console.error('生成回复过程中发生错误:', error);
    showError('生成回复过程中发生错误: ' + error.message);
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
async function sendToGPTBotsAPI(emailContent, customPrompt = null) {
  try {
    // 1. 创建对话
    console.log('创建对话...');
    const conversationResponse = await createConversation();
    if (!conversationResponse.success) {
      return conversationResponse;
    }
    
    currentConversationId = conversationResponse.conversationId;
    console.log('对话创建成功，ID:', currentConversationId);
    
    // 2. 构建消息内容
    const message = customPrompt || `请分析以下邮件内容并提供智能建议：

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
 * 显示翻译结果
 */
function showTranslationResult(emailContent, translationResult) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main" style="padding: 20px;">
      <h2 class="ms-font-xl" style="color: #0078d4; text-align: center;">📝 深度翻译结果</h2>
      
      <!-- 原邮件信息 -->
      <div class="email-info" style="background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <h3 class="ms-font-l">原邮件信息</h3>
        <p><strong>主题:</strong> ${emailContent.subject}</p>
        <p><strong>发件人:</strong> ${emailContent.from}</p>
      </div>
      
      <!-- 翻译结果 -->
      <div class="translation-result" style="background: #fff; border: 1px solid #0078d4; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <h3 class="ms-font-l" style="color: #0078d4;">翻译结果</h3>
        <div style="white-space: pre-wrap; line-height: 1.6;">${translationResult}</div>
      </div>
      
      <div class="action-buttons" style="margin-top: 20px; text-align: center;">
        <div role="button" id="back-to-main-button" class="ms-Button ms-Button--primary" style="margin: 5px;">
          <span class="ms-Button-label">返回主页</span>
        </div>
      </div>
    </div>
  `;
  
  document.getElementById("back-to-main-button").onclick = () => location.reload();
}

/**
 * 显示摘要结果
 */
function showSummaryResult(emailContent, summaryResult) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main" style="padding: 20px;">
      <h2 class="ms-font-xl" style="color: #107c10; text-align: center;">📊 邮件摘要分析</h2>
      
      <!-- 邮件信息 -->
      <div class="email-info" style="background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <h3 class="ms-font-l">邮件信息</h3>
        <p><strong>主题:</strong> ${emailContent.subject}</p>
        <p><strong>发件人:</strong> ${emailContent.from}</p>
        <p><strong>时间:</strong> ${emailContent.dateTimeCreated}</p>
      </div>
      
      <!-- 摘要结果 -->
      <div class="summary-result" style="background: #fff; border: 1px solid #107c10; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <h3 class="ms-font-l" style="color: #107c10;">智能摘要</h3>
        <div style="white-space: pre-wrap; line-height: 1.6;">${summaryResult}</div>
      </div>
      
      <div class="action-buttons" style="margin-top: 20px; text-align: center;">
        <div role="button" id="back-to-main-button" class="ms-Button ms-Button--primary" style="margin: 5px;">
          <span class="ms-Button-label">返回主页</span>
        </div>
      </div>
    </div>
  `;
  
  document.getElementById("back-to-main-button").onclick = () => location.reload();
}

/**
 * 显示回复结果
 */
function showReplyResult(emailContent, replyResult) {
  const appBody = document.getElementById("app-body");
  appBody.innerHTML = `
    <div class="ms-welcome__main" style="padding: 20px;">
      <h2 class="ms-font-xl" style="color: #d83b01; text-align: center;">✍️ 智能回复建议</h2>
      
      <!-- 原邮件信息 -->
      <div class="email-info" style="background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <h3 class="ms-font-l">原邮件信息</h3>
        <p><strong>主题:</strong> ${emailContent.subject}</p>
        <p><strong>发件人:</strong> ${emailContent.from}</p>
        <p><strong>时间:</strong> ${emailContent.dateTimeCreated}</p>
      </div>
      
      <!-- 回复建议 -->
      <div class="reply-suggestion" style="background: #fff; border: 1px solid #d83b01; padding: 15px; border-radius: 5px; margin: 15px 0;">
        <h3 class="ms-font-l" style="color: #d83b01;">回复建议</h3>
        <div style="white-space: pre-wrap; line-height: 1.6;">${replyResult}</div>
      </div>
      
      <div class="action-buttons" style="margin-top: 20px; display: flex; justify-content: center; gap: 10px;">
        <div role="button" id="create-reply-button" class="ms-Button ms-Button--primary">
          <span class="ms-Button-label">创建回复邮件</span>
        </div>
        <div role="button" id="back-to-main-button" class="ms-Button">
          <span class="ms-Button-label">返回主页</span>
        </div>
      </div>
    </div>
  `;
  
  document.getElementById("create-reply-button").onclick = () => createReplyEmail(replyResult);
  document.getElementById("back-to-main-button").onclick = () => location.reload();
}

/**
 * 创建回复邮件
 */
function createReplyEmail(replyContent) {
  try {
    // 提取实际的回复内容（去掉分析部分，只保留回复文本）
    const replyLines = replyContent.split('\n');
    let actualReply = '';
    let foundReplyContent = false;
    
    for (const line of replyLines) {
      if (line.includes('推荐的回复内容') || line.includes('回复内容') || foundReplyContent) {
        foundReplyContent = true;
        if (!line.includes('推荐的回复内容') && !line.includes('：') && line.trim()) {
          actualReply += line + '\n';
        }
      }
    }
    
    // 如果没有找到特定的回复内容，使用完整的结果
    if (!actualReply.trim()) {
      actualReply = replyContent;
    }
    
    // 创建回复邮件
    Office.context.mailbox.item.displayReplyForm(actualReply.trim());
    showSuccess('回复邮件窗口已打开，内容已填入');
  } catch (error) {
    console.error('创建回复邮件时发生错误:', error);
    showError('创建回复邮件时发生错误: ' + error.message);
  }
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
 * 获取操作名称
 */
function getActionName(action) {
  const actionNames = {
    'translate': '深度翻译',
    'summary': '生成摘要', 
    'reply': '生成回复',
    'polish': '内容润色',
    'draft': '生成草稿'
  };
  return actionNames[action] || action;
}

/**
 * 获取当前编辑邮件的内容
 */
async function getCurrentComposeContent() {
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || '');
        } else {
          console.error('获取编辑内容失败:', result.error);
          resolve('');
        }
      });
    } catch (error) {
      console.error('获取编辑内容异常:', error);
      resolve('');
    }
  });
}

/**
 * 获取当前邮件主题
 */
async function getCurrentSubject() {
  try {
    return Office.context.mailbox.item.subject || '新邮件';
  } catch (error) {
    console.error('获取主题失败:', error);
    return '新邮件';
  }
}

/**
 * 获取编辑上下文信息
 */
async function getComposeContext() {
  try {
    const subject = await getCurrentSubject();
    const content = await getCurrentComposeContent();
    return {
      subject: subject,
      content: content
    };
  } catch (error) {
    console.error('获取编辑上下文失败:', error);
    return { subject: '新邮件', content: '' };
  }
}

/**
 * 构建翻译提示词
 */
function buildTranslatePrompt(emailContent) {
  return `请对以下邮件进行深度翻译，保持专业性和语境准确性：

邮件主题: ${emailContent.subject}
发件人: ${emailContent.from}

邮件正文:
${emailContent.body}

请提供：
1. 完整的中文翻译
2. 关键术语解释
3. 语境背景说明（如有必要）`;
}

/**
 * 构建摘要提示词
 */
function buildSummaryPrompt(emailContent) {
  return `请为以下邮件生成详细摘要：

邮件主题: ${emailContent.subject}
发件人: ${emailContent.from}
发送时间: ${emailContent.dateTimeCreated}

邮件正文:
${emailContent.body}

请提供：
1. 邮件核心内容摘要（2-3句话）
2. 关键信息点列表
3. 重要日期和截止时间（如有）
4. 需要采取的行动（如有）
5. 优先级评估`;
}

/**
 * 构建回复提示词
 */
function buildReplyPrompt(emailContent) {
  return `请为以下邮件生成专业的回复内容：

邮件主题: ${emailContent.subject}
发件人: ${emailContent.from}
发送时间: ${emailContent.dateTimeCreated}

邮件正文:
${emailContent.body}

请直接提供完整的回复邮件内容，包括：
1. 推荐的回复内容（完整的邮件内容，包含适当的称呼和结尾）

然后再提供以下分析（用于参考）：
2. 回复要点分析
3. 语气建议（正式/非正式）
4. 需要补充的信息（如有必要）

请确保第1部分的回复内容可以直接复制使用。`;
}

/**
 * 构建编辑模式翻译提示词
 */
function buildComposeTranslatePrompt(content) {
  return `请将以下文本进行专业翻译（根据语言自动识别翻译方向）：

文本内容:
${content}

请提供：
1. 准确的翻译结果
2. 保持原文的语气和风格
3. 适合邮件场景的表达`;
}

/**
 * 构建内容润色提示词
 */
function buildPolishPrompt(content) {
  return `请对以下邮件内容进行润色和优化：

原文内容:
${content}

请提供：
1. 语法和表达的优化
2. 更专业和礼貌的表述
3. 逻辑结构的改善
4. 保持原意的基础上提升质量`;
}

/**
 * 构建编辑模式回复提示词
 */
function buildComposeReplyPrompt(context) {
  return `基于以下信息生成邮件回复内容：

邮件主题: ${context.subject}
当前内容: ${context.content}

请提供：
1. 专业的回复内容建议
2. 适合的开头和结尾
3. 关键要点的回应
4. 商务场景适用的语言`;
}

/**
 * 构建草稿生成提示词
 */
function buildDraftPrompt(subject) {
  return `基于邮件主题生成完整的邮件草稿：

邮件主题: ${subject}

请生成：
1. 合适的邮件开头问候
2. 针对主题的核心内容
3. 专业的结尾和签名建议
4. 整体结构完整、语言得体的邮件内容`;
}

/**
 * 在预览框显示加载状态
 */
function showPreviewLoading(message) {
  const previewContent = document.getElementById('preview-content');
  if (previewContent) {
    previewContent.innerHTML = `
      <div style="text-align: center; color: #0078d4;">
        <div style="display: inline-block; width: 20px; height: 20px; border: 2px solid #f3f3f3; border-top: 2px solid #0078d4; border-radius: 50%; animation: spin 1s linear infinite;"></div>
        <p style="margin-top: 10px;">${message}</p>
      </div>
      <style>
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      </style>
    `;
  }
}

/**
 * 在预览框显示错误
 */
function showPreviewError(message) {
  const previewContent = document.getElementById('preview-content');
  if (previewContent) {
    previewContent.innerHTML = `
      <div style="color: #d13438; text-align: center;">
        <i class="ms-Icon ms-Icon--ErrorBadge" style="font-size: 24px; margin-bottom: 10px;"></i>
        <p>${message}</p>
      </div>
    `;
  }
}

/**
 * 在预览框显示成功消息
 */
function showPreviewSuccess(message) {
  const previewContent = document.getElementById('preview-content');
  if (previewContent) {
    previewContent.innerHTML = `
      <div style="color: #107c10; text-align: center;">
        <i class="ms-Icon ms-Icon--Completed" style="font-size: 24px; margin-bottom: 10px;"></i>
        <p>${message}</p>
      </div>
    `;
    
    // 3秒后恢复正常状态
    setTimeout(() => {
      if (previewContent) {
        clearPreview();
      }
    }, 3000);
  }
}

/**
 * 在预览框显示内容
 */
function showPreviewContent(content, action) {
  const previewContentEl = document.getElementById('preview-content');
  if (previewContentEl) {
    // 存储内容供插入使用
    previewContent = content;
    
    // 格式化显示内容
    let displayContent = content;
    
    // 根据操作类型添加标题
    const actionTitle = getActionName(action);
    displayContent = `【${actionTitle}结果】\n\n${displayContent}`;
    
    // 显示内容
    previewContentEl.textContent = displayContent;
    previewContentEl.style.color = '#323130';
    
    // 显示相应的操作按钮
    const insertBtn = document.getElementById('insert-content-btn');
    if (insertBtn) {
      if (currentMode === 'Compose') {
        // 编辑模式：显示插入按钮
        insertBtn.style.display = 'inline-block';
        insertBtn.textContent = '插入内容';
        insertBtn.onclick = insertContentToEmail;
      } else if (currentMode === 'Read' && action === 'reply') {
        // 阅读模式的回复功能：显示创建回复按钮
        insertBtn.style.display = 'inline-block';
        insertBtn.textContent = '创建回复邮件';
        insertBtn.onclick = createReplyFromPreview;
      } else {
        // 其他情况隐藏按钮
        insertBtn.style.display = 'none';
      }
    }
  }
}

/**
 * 从预览内容创建回复邮件
 */
async function createReplyFromPreview() {
  if (!previewContent || currentMode !== 'Read') {
    console.warn('⚠️ 无预览内容或非阅读模式');
    return;
  }
  
  try {
    showPreviewLoading('正在创建回复邮件...');
    
    // 提取纯净的回复内容
    const cleanReplyContent = extractReplyContent(previewContent);
    
    // 创建回复邮件并填入内容
    Office.context.mailbox.item.displayReplyForm(cleanReplyContent);
    
    showPreviewSuccess('回复邮件已创建，内容已填入编辑器');
    
  } catch (error) {
    console.error('创建回复邮件时发生错误:', error);
    showPreviewError('创建回复邮件时发生错误: ' + error.message);
  }
}

/**
 * 提取纯净的回复内容（去掉分析部分，只保留实际回复文本）
 */
function extractReplyContent(content) {
  try {
    const lines = content.split('\n');
    let replyContent = '';
    let foundReplySection = false;
    let isInReplyContent = false;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      // 寻找回复内容相关的标题
      if (line.includes('推荐的回复内容') || 
          line.includes('回复内容') || 
          line.includes('建议回复') ||
          line.includes('回复建议')) {
        foundReplySection = true;
        isInReplyContent = true;
        continue;
      }
      
      // 如果找到了回复部分，开始收集内容
      if (foundReplySection && isInReplyContent) {
        // 跳过标题行和分隔符
        if (line.startsWith('1.') || 
            line.startsWith('2.') || 
            line.startsWith('3.') || 
            line.startsWith('4.') || 
            line.startsWith('5.') ||
            line.includes('：') || 
            line.includes('要点') ||
            line.includes('分析') ||
            line.includes('建议') ||
            line === '') {
          
          // 如果遇到其他分析项目，停止收集
          if ((line.startsWith('2.') && line.includes('要点')) ||
              (line.startsWith('3.') && line.includes('建议')) ||
              (line.startsWith('4.') && line.includes('信息')) ||
              (line.startsWith('5.') && line.includes('跟进'))) {
            break;
          }
          
          continue;
        }
        
        // 收集实际的回复内容
        if (line && !line.includes('分析') && !line.includes('建议')) {
          replyContent += line + '\n';
        }
      }
    }
    
    // 如果没有找到特定的回复内容结构，尝试智能提取
    if (!replyContent.trim()) {
      replyContent = smartExtractReplyContent(content);
    }
    
    // 最终清理和格式化
    replyContent = replyContent.trim();
    
    // 如果仍然没有内容，使用完整的AI回复
    if (!replyContent) {
      replyContent = content;
    }
    
    // 转换为HTML格式
    replyContent = replyContent.replace(/\n/g, '<br>');
    
    return replyContent;
    
  } catch (error) {
    console.error('提取回复内容失败:', error);
    // 备用方案：使用完整内容
    return content.replace(/\n/g, '<br>');
  }
}

/**
 * 智能提取回复内容（备用方案）
 */
function smartExtractReplyContent(content) {
  try {
    // 寻找常见的回复开头
    const replyStarters = [
      '亲爱的',
      '尊敬的', 
      '您好',
      'Dear',
      'Hi',
      'Hello',
      '感谢您的',
      '谢谢您'
    ];
    
    const lines = content.split('\n');
    let startIndex = -1;
    let endIndex = lines.length;
    
    // 找到回复开始的位置
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      for (const starter of replyStarters) {
        if (line.startsWith(starter)) {
          startIndex = i;
          break;
        }
      }
      if (startIndex !== -1) break;
    }
    
    // 找到回复结束的位置（遇到分析内容）
    for (let i = startIndex + 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line.includes('分析') || 
          line.includes('建议') || 
          line.includes('要点') ||
          line.includes('语气') ||
          line.includes('补充') ||
          line.includes('跟进')) {
        endIndex = i;
        break;
      }
    }
    
    // 提取回复内容
    if (startIndex !== -1) {
      return lines.slice(startIndex, endIndex).join('\n').trim();
    }
    
    // 如果找不到明确的回复结构，返回前半部分内容
    const halfLength = Math.floor(lines.length / 2);
    return lines.slice(0, halfLength).join('\n').trim();
    
  } catch (error) {
    console.error('智能提取回复内容失败:', error);
    return content;
  }
}

// 导出函数以供外部使用
window.deepTranslate = deepTranslate;
window.generateSummary = generateSummary;
window.generateReply = generateReply;
window.toggleDebugInfo = toggleDebugInfo;
window.clearPreview = clearPreview;
window.insertContentToEmail = insertContentToEmail;
window.createReplyFromPreview = createReplyFromPreview;

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