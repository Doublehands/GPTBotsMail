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
    
    // 请求头配置
    headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer app-TdfestItJNTTEMBFnGGBm0Yn'
    },
    
    // 用户ID
    userId: 'word-gpt-plus-user'
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

// 获取创建对话的完整URL（使用代理解决CORS）
function getCreateConversationUrl() {
    const originalUrl = API_CONFIG.baseUrl + API_CONFIG.createConversationEndpoint;
    // 使用cors.io代理解决CORS问题（支持所有HTTP头）
    return `https://cors.io/?${originalUrl}`;
}

// 获取发送消息的完整URL（使用代理解决CORS）
function getChatUrl() {
    const originalUrl = API_CONFIG.baseUrl + API_CONFIG.chatEndpoint;
    // 使用cors.io代理解决CORS问题（支持所有HTTP头）
    return `https://cors.io/?${originalUrl}`;
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