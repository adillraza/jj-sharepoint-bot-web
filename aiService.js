// aiService.js - AI-powered insights and summaries
const axios = require('axios');

class AIService {
    constructor() {
        // For production, these would come from environment variables
        this.endpoint = process.env.AZURE_OPENAI_ENDPOINT || null;
        this.apiKey = process.env.AZURE_OPENAI_API_KEY || null;
        this.deploymentName = process.env.AZURE_OPENAI_DEPLOYMENT_NAME || 'gpt-35-turbo';
        this.apiVersion = '2024-02-15-preview';
        
        // Fallback to a local AI implementation if Azure OpenAI is not configured
        this.useLocalAI = !this.endpoint || !this.apiKey;
        
        if (this.useLocalAI) {
            console.log('🤖 Using local AI implementation (Azure OpenAI not configured)');
        } else {
            console.log('🤖 Using Azure OpenAI for AI insights');
        }
    }

    async generateSummary(documentContent, documentName) {
        if (this.useLocalAI) {
            return this.generateLocalSummary(documentContent, documentName);
        }
        
        try {
            const prompt = `Please provide a concise summary of the following document content from "${documentName}":\n\n${documentContent.substring(0, 3000)}`;
            
            const response = await this.callAzureOpenAI(prompt, 'summary');
            return {
                summary: response,
                source: 'Azure OpenAI',
                confidence: 0.9
            };
        } catch (error) {
            console.error('Azure OpenAI summary failed, falling back to local:', error.message);
            return this.generateLocalSummary(documentContent, documentName);
        }
    }

    async generateInsights(documentContent, documentName) {
        if (this.useLocalAI) {
            return this.generateLocalInsights(documentContent, documentName);
        }

        try {
            const prompt = `Analyze the following document content from "${documentName}" and provide key insights, important points, and actionable items:\n\n${documentContent.substring(0, 3000)}`;
            
            const response = await this.callAzureOpenAI(prompt, 'insights');
            return {
                insights: response,
                source: 'Azure OpenAI',
                confidence: 0.9
            };
        } catch (error) {
            console.error('Azure OpenAI insights failed, falling back to local:', error.message);
            return this.generateLocalInsights(documentContent, documentName);
        }
    }

    async answerQuestion(question, documentContent, documentName) {
        if (this.useLocalAI) {
            return this.answerQuestionLocal(question, documentContent, documentName);
        }

        try {
            const prompt = `Based on the following document content from "${documentName}", please answer this question: "${question}"\n\nDocument content:\n${documentContent.substring(0, 3000)}\n\nAnswer:`;
            
            const response = await this.callAzureOpenAI(prompt, 'qa');
            return {
                answer: response,
                source: 'Azure OpenAI',
                confidence: 0.9,
                documentName: documentName
            };
        } catch (error) {
            console.error('Azure OpenAI Q&A failed, falling back to local:', error.message);
            return this.answerQuestionLocal(question, documentContent, documentName);
        }
    }

    async callAzureOpenAI(prompt, type) {
        const url = `${this.endpoint}/openai/deployments/${this.deploymentName}/chat/completions?api-version=${this.apiVersion}`;
        
        const requestBody = {
            messages: [
                {
                    role: "system",
                    content: "You are a helpful AI assistant that analyzes business documents and provides clear, concise, and actionable insights."
                },
                {
                    role: "user",
                    content: prompt
                }
            ],
            max_tokens: 500,
            temperature: 0.3,
            top_p: 1.0,
            frequency_penalty: 0.0,
            presence_penalty: 0.0
        };

        const response = await axios.post(url, requestBody, {
            headers: {
                'Content-Type': 'application/json',
                'api-key': this.apiKey
            },
            timeout: 30000
        });

        return response.data.choices[0].message.content.trim();
    }

    // Local AI implementation (rule-based for now)
    generateLocalSummary(documentContent, documentName) {
        const sentences = documentContent.split(/[.!?]+/).filter(s => s.trim().length > 20);
        const summary = sentences.slice(0, 3).join('. ') + '.';
        
        return {
            summary: `📄 **Document Summary for ${documentName}:**\n\n${summary}\n\n*This is a basic summary. Configure Azure OpenAI for advanced AI summaries.*`,
            source: 'Local AI',
            confidence: 0.6
        };
    }

    generateLocalInsights(documentContent, documentName) {
        const insights = [];
        const lowerContent = documentContent.toLowerCase();
        
        // Extract key patterns
        const dates = documentContent.match(/\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b|\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2},?\s+\d{4}\b/gi);
        const money = documentContent.match(/\$[\d,]+\.?\d*|\b\d+\.\d{2}\b|\b\d{1,3}(,\d{3})*\s*(dollars?|usd|aud)\b/gi);
        const emails = documentContent.match(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g);
        const phones = documentContent.match(/\b\d{3}[-.]?\d{3}[-.]?\d{4}\b/g);
        
        if (dates) insights.push(`📅 **Important Dates:** ${dates.slice(0, 3).join(', ')}`);
        if (money) insights.push(`💰 **Financial Information:** ${money.slice(0, 3).join(', ')}`);
        if (emails) insights.push(`📧 **Contact Emails:** ${emails.slice(0, 3).join(', ')}`);
        if (phones) insights.push(`📞 **Phone Numbers:** ${phones.slice(0, 3).join(', ')}`);
        
        // Look for action items
        const actionWords = ['must', 'should', 'need to', 'required', 'deadline', 'due', 'urgent', 'important'];
        const actionSentences = documentContent.split(/[.!?]+/).filter(sentence => 
            actionWords.some(word => sentence.toLowerCase().includes(word))
        );
        
        if (actionSentences.length > 0) {
            insights.push(`⚡ **Action Items:** ${actionSentences.slice(0, 2).join('. ')}`);
        }
        
        return {
            insights: insights.length > 0 ? insights.join('\n\n') : `📋 **Key Points from ${documentName}:**\n\nThis document contains important business information. Configure Azure OpenAI for detailed AI insights.`,
            source: 'Local AI',
            confidence: 0.7
        };
    }

    answerQuestionLocal(question, documentContent, documentName) {
        // This is a simplified version - the DocumentProcessor already handles this better
        const lowerQuestion = question.toLowerCase();
        const lowerContent = documentContent.toLowerCase();
        
        // Find relevant sentences
        const sentences = documentContent.split(/[.!?]+/).filter(s => s.trim().length > 10);
        const questionWords = lowerQuestion.split(' ').filter(word => word.length > 3);
        
        const relevantSentences = sentences.filter(sentence => 
            questionWords.some(word => sentence.toLowerCase().includes(word))
        );
        
        if (relevantSentences.length > 0) {
            return {
                answer: `Based on ${documentName}: ${relevantSentences.slice(0, 2).join(' ')}`,
                source: 'Local AI',
                confidence: 0.6,
                documentName: documentName
            };
        }
        
        return {
            answer: `I found ${documentName} but couldn't locate specific information about "${question}". Configure Azure OpenAI for better AI-powered answers.`,
            source: 'Local AI',
            confidence: 0.3,
            documentName: documentName
        };
    }
}

module.exports = { AIService };
