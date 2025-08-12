/*
 * GPTBots API配置文件 - 简化版本
 * 只保留最基本的API调用配置
 */

// API配置对象
const API_CONFIG = {
    // 多个代理URL，按优先级排序
    proxyUrls: [
        'https://api-sg.gptbots.ai',  // 直接尝试（可能因CORS失败）
        'https://cors-anywhere.herokuapp.com/https://api-sg.gptbots.ai',
        'https://corsproxy.io/?https://api-sg.gptbots.ai'
    ],
    
    // 当前使用的代理索引
    currentProxyIndex: 0,
    
    // 创建对话端点
    createConversationEndpoint: '/v1/conversation',
    
    // 发送消息端点
    chatEndpoint: '/v2/conversation/message',
    
    // 不同技能的API密钥配置
    apiKeys: {
        translate: 'app-6GQY5ONwN73Spp7Li9Bz8o37',    // 深度翻译
        summary: 'app-BHxaWqTPqQiyein42aVWqkDO',     // 生成摘要
        reply: 'app-TdfestItJNTTEMBFnGGBm0Yn'        // 生成回复
    },
    
    // 请求头配置（基础模板）
    getHeaders: function(skillType) {
        return {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.apiKeys[skillType] || this.apiKeys.reply}`
        };
    },
    
    // 用户ID
    userId: 'outlook-addin-user'
};

// 构建创建对话的请求数据
function buildCreateConversationData() {
    return {
        user_id: API_CONFIG.userId
    };
}

// 构建发送消息的请求数据
function buildChatMessageData(conversationId, message) {
    return {
        conversation_id: conversationId,
        response_mode: 'blocking',
        messages: [
            {
                role: "user",
                content: message
            }
        ]
    };
}

// 获取当前代理的基础URL
function getCurrentProxyUrl() {
    return API_CONFIG.proxyUrls[API_CONFIG.currentProxyIndex];
}

// 获取创建对话的完整URL
function getCreateConversationUrl() {
    return getCurrentProxyUrl() + API_CONFIG.createConversationEndpoint;
}

// 获取发送消息的完整URL
function getChatUrl() {
    return getCurrentProxyUrl() + API_CONFIG.chatEndpoint;
}

// 切换到下一个代理
function switchToNextProxy() {
    API_CONFIG.currentProxyIndex = (API_CONFIG.currentProxyIndex + 1) % API_CONFIG.proxyUrls.length;
    console.log(`🔄 切换到代理 ${API_CONFIG.currentProxyIndex + 1}:`, getCurrentProxyUrl());
}

// 重置到第一个代理
function resetProxy() {
    API_CONFIG.currentProxyIndex = 0;
}

// 解析创建对话的响应
function parseCreateConversationResponse(response) {
    return {
        success: true,
        conversationId: response.conversation_id,
        data: response
    };
}

// 解析聊天消息的响应
function parseChatResponse(response) {
    return {
        success: true,
        answer: response.output[0].content.text,
        data: response
    };
}