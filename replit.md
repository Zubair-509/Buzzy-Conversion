# PDF to DOCX Converter

## Overview

A Flask-based web application that converts PDF documents to DOCX format while preserving formatting. The application provides a user-friendly drag-and-drop interface for file uploads, validates PDF files, performs conversion using the pdf2docx library, and allows users to download the converted documents.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture
- **Template Engine**: Jinja2 templates with Bootstrap 5 dark theme for responsive UI
- **Interactive Components**: JavaScript-based drag-and-drop file upload with real-time validation
- **Styling**: Custom CSS with Bootstrap integration for consistent dark theme appearance
- **User Experience**: Progress indicators, file validation feedback, and download management

### Backend Architecture  
- **Web Framework**: Flask with standard routing and request handling
- **File Processing**: pdf2docx library for PDF to DOCX conversion with PyPDF2 for validation
- **File Management**: Separate upload and converted directories with secure filename handling
- **Error Handling**: Comprehensive validation for file types, sizes, and PDF integrity
- **Security**: File extension validation, secure filename generation, and request size limits

### Data Storage
- **File Storage**: Local filesystem with organized directories for uploads and converted files
- **Temporary Processing**: Uses system temporary directories during conversion process
- **Session Management**: Flask sessions with configurable secret keys for security

### Configuration Management
- **Environment Variables**: SESSION_SECRET for production security
- **File Limits**: 50MB maximum file size with configurable upload restrictions
- **Directory Structure**: Automatic creation of required upload and conversion directories

## External Dependencies

### Core Libraries
- **Flask**: Web framework for routing, templating, and request handling
- **pdf2docx**: Primary library for PDF to DOCX conversion functionality
- **PyPDF2**: PDF validation and integrity checking
- **Werkzeug**: Secure filename handling and HTTP utilities

### Frontend Dependencies
- **Bootstrap 5**: UI framework with Replit dark theme integration
- **Font Awesome 6.4.0**: Icon library for enhanced user interface elements

### System Requirements
- **Python Environment**: Flask-compatible Python runtime
- **File System**: Local storage for temporary file processing and conversion