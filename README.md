# Overview

This is a Streamlit-based web application that serves as an AI-powered editing assistant for Microsoft Word documents stored in OneDrive. The application allows users to authenticate with their Microsoft 365 account, browse Word documents from OneDrive, and use natural language commands to edit documents with the help of Google's Gemini AI. The edited documents can be saved back to OneDrive, creating a seamless document editing workflow.

# User Preferences

Preferred communication style: Simple, everyday language.

# System Architecture

## Frontend Architecture

**Streamlit Web Application**
- The application uses Streamlit as the web framework for rapid UI development and interactive components
- Session state management maintains conversation history, current document selection, and document content across user interactions
- Resource caching is implemented via `@st.cache_resource` decorator to persist service instances (OneDrive client, document processor, LLM service) across reruns
- The UI is organized with a sidebar for document selection and main area for chat-based interaction

**Rationale**: Streamlit was chosen for its ability to quickly build data-focused web applications with minimal frontend code, making it ideal for proof-of-concept AI tools. The declarative nature allows developers to focus on application logic rather than UI implementation details.

## Backend Architecture

**Service-Oriented Design**
- The application follows a service-oriented architecture with three core services:
  1. `OneDriveClient`: Handles Microsoft Graph API integration
  2. `DocumentProcessor`: Manages Word document parsing and creation
  3. `LLMService`: Interfaces with Google Gemini AI for document editing

- Each service is independently initialized and can be tested/modified in isolation
- Asynchronous operations are used for OneDrive API calls to improve performance

**Rationale**: Separating concerns into distinct services improves maintainability and testability. The async pattern for OneDrive operations prevents blocking the UI during network requests.

## Document Processing Strategy

**python-docx Library**
- Documents are processed using the `python-docx` library for reading and writing .docx files
- Text extraction includes both paragraphs and table content, preserving structure with separator notation (` | `) for table cells
- Documents are handled as byte streams to avoid filesystem dependencies

**Rationale**: python-docx provides reliable Word document manipulation without requiring Microsoft Office installation. Byte stream handling enables seamless integration with cloud storage APIs.

## LLM Integration Pattern

**Google Gemini API with Structured Outputs**
- The application uses Google's Gemini 2.5 Flash model for document editing
- System prompts guide the LLM to act as a document editor with specific formatting requirements
- Responses are structured as JSON with fields for edited content, explanation, and change summary
- The model is instructed to preserve document structure while applying user-requested changes

**Rationale**: Gemini provides strong language understanding for interpreting edit requests. JSON-structured responses ensure consistent parsing and enable the application to separate edited content from explanatory text.

## Authentication and Authorization

**Microsoft Graph API via Replit Connectors**
- Authentication is handled through Replit's connector system for OneDrive/Microsoft 365
- Access tokens are retrieved from Replit's connector API using environment-based identity tokens
- Token caching is implemented to reduce API calls, with validation to refresh when expired
- The application uses either REPL_IDENTITY (for development) or WEB_REPL_RENEWAL (for deployed apps) for authentication

**Rationale**: Replit connectors abstract OAuth2 flow complexity, allowing developers to focus on application logic. The dual token approach supports both development and production environments.

# External Dependencies

## Cloud Storage
- **Microsoft OneDrive** (via Microsoft Graph API v1.0): Primary document storage and retrieval
  - Used for listing Word documents
  - Downloading document content
  - Uploading modified documents
  - Requires connector configuration in Replit environment

## AI/ML Services
- **Google Gemini API** (gemini-2.5-flash model): Natural language processing for document editing
  - Requires GEMINI_API_KEY environment variable
  - Processes edit requests and generates modified document content
  - Provides explanations of changes made

## Python Libraries
- **streamlit**: Web application framework and UI components
- **python-docx**: Word document parsing and creation
- **aiohttp**: Asynchronous HTTP client for OneDrive API calls
- **google-genai**: Google Gemini AI SDK

## Environment Configuration
Required environment variables:
- `GEMINI_API_KEY`: API key for Google Gemini service
- `REPLIT_CONNECTORS_HOSTNAME`: Hostname for Replit connector API
- `REPL_IDENTITY` or `WEB_REPL_RENEWAL`: Authentication tokens for Replit environment
