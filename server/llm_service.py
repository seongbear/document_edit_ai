import os
import json
from openai import OpenAI
from typing import Dict, Any, List, Optional

class LLMService:
    """Service for LLM-powered document editing"""
    
    def __init__(self):
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise Exception("OPENAI_API_KEY environment variable is required")
        
        self.client = OpenAI(api_key=api_key)
        # the newest OpenAI model is "gpt-5" which was released August 7, 2025.
        # do not change this unless explicitly requested by the user
        self.model = "gpt-5"
    
    def process_edit_request(self, document_content: str, user_request: str) -> Dict[str, Any]:
        """Process a user's edit request for a document"""
        try:
            system_prompt = """You are an expert document editor and writing assistant. Your task is to help users edit and improve their Microsoft Word documents based on their specific requests.

When given a document and an edit request:
1. Analyze the user's request carefully
2. Apply the requested changes to the document content
3. Maintain the document's overall structure and formatting intent
4. Provide a clear explanation of what changes you made

Respond with JSON in this exact format:
{
    "edited_content": "The complete edited document text",
    "explanation": "A clear explanation of the changes made",
    "changes_summary": "Brief summary of key changes"
}

Important guidelines:
- Keep the same paragraph structure unless specifically asked to change it
- Preserve important information while making requested improvements
- If the request is unclear, make reasonable assumptions and explain them
- For table-like content with | separators, maintain that format
- Always provide the complete edited document, not just the changes"""

            user_prompt = f"""Document Content:
{document_content}

User Request: {user_request}

Please edit the document according to the user's request and respond with the JSON format specified."""

            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                response_format={"type": "json_object"}
            )
            
            content = response.choices[0].message.content
            if not content:
                raise Exception("Empty response from LLM")
            result = json.loads(content)
            
            # Validate response structure
            if not all(key in result for key in ["edited_content", "explanation"]):
                raise Exception("Invalid response format from LLM")
            
            return result
            
        except json.JSONDecodeError as e:
            raise Exception(f"Failed to parse LLM response: {str(e)}")
        except Exception as e:
            raise Exception(f"Failed to process edit request: {str(e)}")
    
    def analyze_document(self, document_content: str) -> Dict[str, Any]:
        """Analyze document and provide insights"""
        try:
            system_prompt = """You are a document analysis expert. Analyze the given document and provide insights about its structure, content, and potential improvements.

Respond with JSON in this format:
{
    "word_count": estimated_word_count,
    "document_type": "type of document (e.g., report, letter, article)",
    "tone": "writing tone (e.g., formal, casual, academic)",
    "structure_analysis": "analysis of document structure",
    "improvement_suggestions": ["list", "of", "improvement", "suggestions"],
    "key_topics": ["main", "topics", "covered"]
}"""

            user_prompt = f"Please analyze this document:\n\n{document_content}"

            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                response_format={"type": "json_object"}
            )
            
            content = response.choices[0].message.content
            if not content:
                raise Exception("Empty response from LLM")
            return json.loads(content)
            
        except Exception as e:
            raise Exception(f"Failed to analyze document: {str(e)}")
    
    def suggest_improvements(self, document_content: str) -> List[str]:
        """Generate improvement suggestions for the document"""
        try:
            analysis = self.analyze_document(document_content)
            return analysis.get("improvement_suggestions", [])
            
        except Exception as e:
            raise Exception(f"Failed to generate suggestions: {str(e)}")
    
    def fix_grammar_and_style(self, document_content: str) -> Dict[str, Any]:
        """Fix grammar and improve writing style"""
        return self.process_edit_request(
            document_content,
            "Please fix any grammar mistakes and improve the writing style while keeping the same meaning and structure."
        )
    
    def summarize_document(self, document_content: str, target_length: str = "medium") -> str:
        """Create a summary of the document"""
        try:
            length_instructions = {
                "short": "in 2-3 sentences",
                "medium": "in 1-2 paragraphs",
                "long": "in 3-4 paragraphs with key details"
            }
            
            instruction = length_instructions.get(target_length, length_instructions["medium"])
            
            prompt = f"Please summarize the following document {instruction}:\n\n{document_content}"
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}]
            )
            
            content = response.choices[0].message.content
            return content if content else ""
            
        except Exception as e:
            raise Exception(f"Failed to summarize document: {str(e)}")
    
    def generate_chat_response(self, message: str, context: Optional[str] = None) -> str:
        """Generate a general chat response"""
        try:
            system_prompt = "You are a helpful assistant for document editing. Provide clear, concise responses to user questions about document editing, writing, and content improvement."
            
            if context:
                user_prompt = f"Context: {context}\n\nUser question: {message}"
            else:
                user_prompt = message
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ]
            )
            
            content = response.choices[0].message.content
            return content if content else ""
            
        except Exception as e:
            raise Exception(f"Failed to generate chat response: {str(e)}")
