// documentProcessor.js - Document Content Extraction and Q&A
const mammoth = require('mammoth');
const pdf = require('pdf-parse');

class DocumentProcessor {
    constructor() {
        // For now, we'll use a simple text-based Q&A
        // Later we can integrate with Azure OpenAI or other AI services
    }

    async extractTextFromDocument(buffer, mimeType, fileName) {
        try {
            console.log(`📄 Extracting text from ${fileName} (${mimeType})`);
            
            if (mimeType.includes('application/vnd.openxmlformats-officedocument.wordprocessingml.document') || 
                fileName.endsWith('.docx')) {
                return await this.extractFromDocx(buffer);
            } else if (mimeType.includes('application/pdf') || fileName.endsWith('.pdf')) {
                return await this.extractFromPdf(buffer);
            } else if (mimeType.includes('text/') || fileName.endsWith('.txt')) {
                return buffer.toString('utf-8');
            } else {
                throw new Error(`Unsupported file type: ${mimeType}`);
            }
        } catch (error) {
            console.error('❌ Document extraction failed:', error);
            throw error;
        }
    }

    async extractFromDocx(buffer) {
        const result = await mammoth.extractRawText({ buffer });
        return result.value;
    }

    async extractFromPdf(buffer) {
        const data = await pdf(buffer);
        return data.text;
    }

    // Simple keyword-based Q&A (can be enhanced with AI later)
    async answerQuestion(question, documentText, documentName) {
        const lowerQuestion = question.toLowerCase();
        const lowerText = documentText.toLowerCase();
        
        // Simple keyword matching and context extraction
        const sentences = documentText.split(/[.!?]+/).filter(s => s.trim().length > 10);
        const relevantSentences = [];

        // Extract keywords from question
        const questionWords = lowerQuestion
            .replace(/[^\w\s]/g, ' ')
            .split(/\s+/)
            .filter(word => word.length > 3 && !['what', 'when', 'where', 'how', 'why', 'who', 'does', 'will', 'can', 'should'].includes(word));

        console.log(`🔍 Looking for keywords: ${questionWords.join(', ')}`);

        // Find sentences containing question keywords
        for (const sentence of sentences) {
            const lowerSentence = sentence.toLowerCase();
            const matchCount = questionWords.filter(word => lowerSentence.includes(word)).length;
            
            if (matchCount > 0) {
                relevantSentences.push({
                    text: sentence.trim(),
                    relevance: matchCount / questionWords.length
                });
            }
        }

        // Sort by relevance and take top results
        relevantSentences.sort((a, b) => b.relevance - a.relevance);
        const topSentences = relevantSentences.slice(0, 3);

        if (topSentences.length === 0) {
            return {
                answer: "I couldn't find specific information related to your question in this document.",
                confidence: 0,
                sources: []
            };
        }

        // Generate a more conversational answer
        const answer = this.generateConversationalAnswer(question, topSentences, documentName);
        const confidence = topSentences[0].relevance;

        return {
            answer: answer,
            confidence: confidence,
            sources: [`${documentName} (${topSentences.length} relevant passages found)`],
            documentName: documentName
        };
    }

    generateConversationalAnswer(question, topSentences, documentName) {
        const lowerQuestion = question.toLowerCase();
        const context = topSentences.map(s => s.text).join(' ');
        
        // Question type detection for better responses
        if (lowerQuestion.includes('what is') || lowerQuestion.includes('what are')) {
            return `Based on ${documentName}, here's what I found:\n\n${context}`;
        } 
        
        if (lowerQuestion.includes('when') || lowerQuestion.includes('date')) {
            const datePattern = /\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b|\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2},?\s+\d{4}\b/gi;
            const dates = context.match(datePattern);
            if (dates && dates.length > 0) {
                return `📅 **Dates found in ${documentName}:**\n${dates.join(', ')}\n\n**Context:** ${context}`;
            }
        }
        
        if (lowerQuestion.includes('how much') || lowerQuestion.includes('cost') || lowerQuestion.includes('price') || lowerQuestion.includes('budget')) {
            const moneyPattern = /\$[\d,]+\.?\d*|\b\d+\.\d{2}\b|\b\d{1,3}(,\d{3})*\s*(dollars?|usd|aud)\b/gi;
            const money = context.match(moneyPattern);
            if (money && money.length > 0) {
                return `💰 **Financial information from ${documentName}:**\n${money.join(', ')}\n\n**Details:** ${context}`;
            }
        }
        
        if (lowerQuestion.includes('who') || lowerQuestion.includes('person') || lowerQuestion.includes('contact')) {
            const namePattern = /\b[A-Z][a-z]+ [A-Z][a-z]+\b/g;
            const emailPattern = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
            const names = context.match(namePattern) || [];
            const emails = context.match(emailPattern) || [];
            
            if (names.length > 0 || emails.length > 0) {
                let response = `👥 **People mentioned in ${documentName}:**\n`;
                if (names.length > 0) response += `Names: ${names.join(', ')}\n`;
                if (emails.length > 0) response += `Emails: ${emails.join(', ')}\n`;
                response += `\n**Context:** ${context}`;
                return response;
            }
        }
        
        if (lowerQuestion.includes('summary') || lowerQuestion.includes('summarize')) {
            return `📋 **Summary from ${documentName}:**\n\n${context}\n\n*This summary is based on the most relevant sections of the document.*`;
        }
        
        if (lowerQuestion.includes('deadline') || lowerQuestion.includes('due date')) {
            const deadlinePattern = /(deadline|due date|expires?|by)\s+([^.!?]+)/gi;
            const deadlines = context.match(deadlinePattern);
            if (deadlines && deadlines.length > 0) {
                return `⏰ **Deadlines from ${documentName}:**\n${deadlines.join('\n')}\n\n**Full context:** ${context}`;
            }
        }
        
        // Default conversational response
        return `📄 **From ${documentName}:**\n\n${context}\n\n*I found this information that seems relevant to your question. Would you like me to search for anything more specific?*`;
    }

    // Enhanced Q&A with Azure OpenAI (for later implementation)
    async answerQuestionWithAI(question, documentText, documentName) {
        // This would integrate with Azure OpenAI
        // For now, fall back to simple keyword matching
        return await this.answerQuestion(question, documentText, documentName);
    }

    // Extract key information from documents
    extractKeyInfo(documentText, documentName) {
        const info = {
            wordCount: documentText.split(/\s+/).length,
            paragraphs: documentText.split(/\n\s*\n/).length,
            hasNumbers: /\d+/.test(documentText),
            hasDates: /\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}|\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}/.test(documentText),
            hasEmails: /@/.test(documentText),
            documentName: documentName
        };

        // Extract potential key topics (simple approach)
        const words = documentText.toLowerCase()
            .replace(/[^\w\s]/g, ' ')
            .split(/\s+/)
            .filter(word => word.length > 4);
        
        const wordCounts = {};
        words.forEach(word => {
            wordCounts[word] = (wordCounts[word] || 0) + 1;
        });

        info.topKeywords = Object.entries(wordCounts)
            .sort(([,a], [,b]) => b - a)
            .slice(0, 10)
            .map(([word, count]) => ({ word, count }));

        return info;
    }
}

module.exports = { DocumentProcessor };
