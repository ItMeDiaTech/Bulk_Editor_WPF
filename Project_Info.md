# WPF Bulk Editor - Complete Development Prompt

## 🎯 **Application Overview**

Create a professional WPF desktop application called "Bulk Editor" that processes Microsoft Word documents (.docx files) to perform comprehensive hyperlink management and document cleanup operations. The application must handle both individual files and batch folder processing with a modern, responsive user interface and enterprise-grade reliability.

## 📋 **Core Functionality Requirements**

### **Primary Document Processing Features**

- **Hyperlink Validation & Repair**: Detect and fix broken hyperlinks using external API integration
- **Content ID Management**: Automatically append content identifiers to hyperlink text
- **Internal Link Processing**: Validate and repair document internal anchors and cross-references
- **Text Optimization**: Eliminate double spaces and normalize formatting
- **Bulk Hyperlink Replacement**: Apply rule-based find-and-replace operations across multiple links
- **Changelog Generation**: Create detailed processing reports with before/after comparisons

### **User Experience Features**

- **File Selection**: Support both individual file selection and folder-based batch processing
- **Real-time Progress**: Comprehensive progress reporting with cancellation support
- **Settings Management**: Persistent user preferences and processing options
- **Error Handling**: Robust error recovery with detailed logging and user feedback
- **Backup Creation**: Automatic file backups before processing with configurable retention
- **Theme Support**: Modern light/dark theme switching with MaterialDesignInXamlToolkit with possible other themes

### **Enterprise Features**

- **Comprehensive Logging**: Detailed audit trails for all operations
- **Configuration Management**: JSON-based settings with validation
- **Performance Monitoring**: Processing statistics and performance metrics
- **Safe File Operations**: Atomic file processing with rollback capabilities

## 🛠️ **Required Technology Stack (Free Frameworks Only)**

### **Core Framework Requirements**

- **.NET 8 LTS**: Target framework for long-term support through November 2026
- **WPF**: Windows Presentation Foundation for the user interface
- **CommunityToolkit.Mvvm**: Microsoft's official MVVM framework with source generators for clean, maintainable architecture

### **Architecture & Infrastructure**

- **Microsoft.Extensions.DependencyInjection**: Built-in dependency injection container for service management
- **Microsoft.Extensions.Configuration**: Modern configuration management with JSON support
- **Microsoft.Extensions.Hosting**: Host builder pattern for proper application lifecycle management
- **Microsoft.Extensions.Http**: HTTP client factory for API communications

### **Document Processing**

- **DocumentFormat.OpenXml**: Microsoft's official library for Word document manipulation
- **System.Text.Json**: Built-in JSON processing for configuration and API responses

### **Logging & Monitoring**

- **Serilog**: Industry-standard structured logging framework
- **Serilog.Sinks.Console**: Console output for development
- **Serilog.Sinks.File**: File-based logging with rotation support
- **Serilog.Extensions.Logging**: Integration with Microsoft.Extensions.Logging

### **UI & User Experience**

- **MaterialDesignInXamlToolkit**: Professional Material Design implementation for WPF
- **MaterialDesignThemes**: Theming support with light/dark modes
- **MaterialDesignColors**: Comprehensive color palette management

## 🏗️ **Architecture Requirements**

### **Design Patterns & Architecture**

- **Clean Architecture**: Proper separation of concerns with distinct layers
- **MVVM Pattern**: Complete Model-View-ViewModel implementation using CommunityToolkit.Mvvm
- **Service Layer Pattern**: Business logic abstraction through service interfaces
- **Repository Pattern**: Data access abstraction for settings and file operations
- **Command Pattern**: User interactions handled through ICommand implementations
- **Observer Pattern**: Progress reporting and event-driven communication

### **Async Programming Requirements**

- **Task-based Asynchronous Pattern (TAP)**: All I/O operations must be fully asynchronous
- **IProgress<T> Implementation**: Structured progress reporting for long-running operations
- **CancellationToken Support**: Proper cancellation handling throughout the application
- **UI Thread Management**: Ensure UI responsiveness during background processing

### **Error Handling & Reliability**

- **Global Exception Handling**: Application-level exception management
- **Structured Error Reporting**: User-friendly error messages with technical details in logs
- **Retry Mechanisms**: Configurable retry policies for network operations
- **Graceful Degradation**: Application continues functioning when non-critical features fail

## 📁 **Project Structure Requirements**

### **Solution Organization**

- **Views/**: XAML user interfaces with proper data binding
- **ViewModels/**: Business logic and UI state management using ObservableObject
- **Services/**: Business service implementations with dependency injection
- **Services/Abstractions/**: Service interface definitions
- **Models/**: Data models and DTOs for application state
- **Configuration/**: Settings models and configuration management
- **Extensions/**: Extension methods and utility functions
- **Resources/**: Themes, styles, and localized resources
- **Converters/**: Value converters for XAML data binding

## ⚙️ **Technical Implementation Requirements**

### **Configuration Management**

- **appsettings.json**: Primary configuration file with strongly-typed models
- **User Settings**: Separate user preferences with persistent storage
- **Environment Support**: Development/Production configuration variations
- **Settings Validation**: Data annotations for configuration validation

### **Logging Implementation**

- **Structured Logging**: JSON-formatted log entries with contextual data
- **Log Levels**: Appropriate use of Debug, Information, Warning, Error levels
- **Performance Logging**: Operation timing and performance metrics
- **User Action Tracking**: Audit trail for user interactions
- **File Rotation**: Automatic log file management with size and date-based rotation

### **Background Processing**

- **Producer-Consumer Pattern**: Queue-based processing for multiple documents
- **Chunked Processing**: Handle large documents without memory issues
- **Progress Aggregation**: Combine progress from multiple operations
- **Resource Management**: Proper disposal of document resources

### **File System Operations**

- **Atomic Operations**: Write-then-rename pattern for safe file updates
- **Backup Strategies**: Automatic backup creation with configurable retention
- **Path Validation**: Secure path handling to prevent directory traversal
- **Concurrent Access**: Handle file locking and sharing violations

## 🎨 **User Interface Requirements**

### **Material Design Implementation**

- **Modern Aesthetics**: Clean, professional interface following Material Design principles
- **Responsive Layout**: Adaptive layouts that work across different window sizes
- **Accessibility**: Proper contrast ratios, keyboard navigation, and screen reader support
- **Animation**: Smooth transitions and micro-interactions for enhanced user experience

### **User Experience Features**

- **Drag & Drop**: File and folder drag-drop support
- **Keyboard Shortcuts**: Comprehensive keyboard navigation and shortcuts
- **Context Menus**: Right-click actions for common operations
- **Status Indicators**: Clear visual feedback for application state
- **Tool Tips**: Helpful descriptions for all interactive elements

### **Theme Management**

- **System Integration**: Automatic theme detection from Windows settings
- **Theme Persistence**: Remember user theme preferences
- **Dynamic Switching**: Runtime theme changes without application restart
- **Custom Styling**: Consistent styling across all UI elements

## 🔄 **Development Approach Instructions**

### **Phase 1: Planning & Architecture**

Before writing any implementation, create a comprehensive development plan that addresses:

1. **Project Structure Design**

   - Detailed folder organization with rationale
   - Service dependency mapping and injection strategy
   - MVVM implementation approach using CommunityToolkit.Mvvm
   - Data flow architecture between layers

2. **Technical Architecture Planning**

   - Service interface definitions and implementations
   - Configuration model design with validation requirements
   - Error handling strategy across all layers
   - Logging implementation with structured data design

3. **User Interface Design**

   - Main window layout with Material Design integration
   - View-ViewModel binding strategy
   - Progress reporting UI implementation
   - Settings dialog design and data binding

4. **Implementation Roadmap**
   - Development phases with clear milestones
   - Testing strategy for each component
   - Integration points between services
   - Deployment and packaging considerations

### **Phase 2: Implementation Strategy**

After completing the planning phase, implement the solution following these guidelines:

1. **Start with Foundation**

   - Project setup with all required NuGet packages
   - Dependency injection configuration
   - Basic MVVM infrastructure with CommunityToolkit.Mvvm
   - Configuration management implementation

2. **Build Core Services**

   - Document processing service with OpenXML integration
   - Logging service implementation with Serilog
   - Settings service with JSON persistence
   - HTTP service for API communications

3. **Develop User Interface**

   - Main window with Material Design theming
   - Progress reporting implementation
   - Settings management interface
   - Error handling and user feedback systems

4. **Integration & Testing**
   - Service integration testing
   - End-to-end functionality validation
   - Error scenario testing
   - Performance optimization

## ✅ **Quality Standards**

### **Best Practices Requirements**

- **SOLID Principles**: Follow all SOLID design principles throughout the codebase
- **DRY Implementation**: Eliminate code duplication through proper abstraction
- **Testability**: Design all components for easy unit and integration testing
- **Documentation**: Comprehensive XML documentation for all public APIs
- **Error Messages**: User-friendly error messages with actionable guidance

### **Performance Expectations**

- **Responsive UI**: UI remains responsive during all operations
- **Memory Management**: Proper disposal of resources and memory cleanup
- **Scalability**: Handle large documents and batch operations efficiently
- **Startup Time**: Fast application startup with lazy loading where appropriate

### **Security Considerations**

- **Input Validation**: Validate all user inputs and file paths
- **Safe File Operations**: Prevent unauthorized file system access
- **API Security**: Secure handling of API keys and network communications
- **Error Information**: Avoid exposing sensitive information in error messages

## 🚀 **Success Criteria**

The completed application must demonstrate:

- Professional-grade user interface with Material Design implementation
- Robust document processing capabilities with comprehensive error handling
- Modern asynchronous architecture with proper progress reporting
- Enterprise-quality logging and configuration management
- Extensible design allowing for future feature additions
- Comprehensive testing coverage ensuring reliability

## 📝 **Deliverable Instructions**

1. **Begin with detailed planning** - Do not start coding until you have created a comprehensive architectural plan
2. **Ask clarifying questions** - Request clarification on any ambiguous requirements before implementation
3. **Follow the technology stack exactly** - Use only the specified free frameworks and libraries
4. **Implement incrementally** - Build and test each component before moving to the next
5. **Maintain clean architecture** - Ensure proper separation of concerns throughout development
6. **Document decisions** - Explain architectural and implementation choices as you progress

Create a production-ready application that demonstrates modern WPF development best practices while maintaining simplicity and reliability in its core functionality.
