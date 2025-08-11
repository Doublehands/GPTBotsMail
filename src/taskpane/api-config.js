/*
 * GPTBots API配置文件 - 简化版本
 * 只保留最基本的API调用配置
 */

// API配置对象
const API_CONFIG = {
    // GPTBots API基础URL
    baseUrl: 'https://api.gptbots.ai',
    
    // 创建对话端点
    createConversationEndpoint: '/v1/conversation',
    
    // 发送消息端点
    chatEndpoint: '/v2/conversation/message',
    
    // 不同技能的API密钥配置
    // 暂时都使用已验证可用的API密钥
    apiKeys: {
        translate: 'app-TdfestItJNTTEMBFnGGBm0Yn',    // 深度翻译 (暂用回复密钥)
        summary: 'app-TdfestItJNTTEMBFnGGBm0Yn',     // 生成摘要 (暂用回复密钥)
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
        inputs: {},
        query: message,
        response_mode: 'blocking',
        user: API_CONFIG.userId
    };
}

// 获取创建对话的完整URL（直接调用，不使用代理）
function getCreateConversationUrl() {
    return API_CONFIG.baseUrl + API_CONFIG.createConversationEndpoint;
}

// 获取发送消息的完整URL（直接调用，不使用代理）
function getChatUrl() {
    return API_CONFIG.baseUrl + API_CONFIG.chatEndpoint;
}

// 解析创建对话的响应
function parseCreateConversationResponse(response) {
    return {
        success: true,
        conversationId: response.data.conversation_id,
        data: response
    };
}

// 解析聊天消息的响应
function parseChatResponse(response) {
    return {
        success: true,
        answer: response.data.answer,
        data: response
    };
}