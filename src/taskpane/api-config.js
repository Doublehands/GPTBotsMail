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

// 模拟API响应数据
const MOCK_RESPONSES = {
    translate: `深度翻译：
主题：关于GPTBots平台AI电商客服解决方案的咨询

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
Jacky`,
    summary: `询问邮件重点：

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

如需优先解决特定痛点（如退换货自动化），可提供更详细业务场景供GPTBots优化配置。`,
    reply: `Dear Jacky,

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
contact@gptbots.ai`
};

// 获取创建对话的完整URL（模拟）
function getCreateConversationUrl() {
    return 'mock://conversation';
}

// 获取发送消息的完整URL（模拟）
function getChatUrl() {
    return 'mock://chat';
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