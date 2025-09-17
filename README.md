# Learning Support System

A collection of ASP.NET VB web services designed to integrate a web portal with SCORM course players using content served from MinIO object storage.

## Overview

This system provides a comprehensive learning management platform that handles user authentication, course access, SCORM content delivery, assessments, and document management. The architecture consists of multiple web services that work together to provide a seamless e-learning experience.

## Architecture

### Core Components

- **Web Services**: ASP.NET VB HTTP handlers (.ashx files) that provide REST-like API endpoints
- **Database Layer**: SQL Server databases for user management, course tracking, and content metadata
- **Content Storage**: MinIO object storage for SCORM packages and course materials
- **Frontend**: HTML/JavaScript interfaces for course access and management

### Database Structure

The system uses four main SQL Server databases:

1. **siebeldb**: Main application database containing user accounts, subscriptions, courses, and registrations
2. **elearning**: E-learning specific data including SCORM tracking, KBA questions, and player data
3. **DMS**: Document Management System for course materials and attachments
4. **reports**: Logging and reporting data

## Web Services

### Authentication & Session Management

#### WsLogin.ashx
- **Purpose**: Primary login service for user authentication
- **Parameters**: 
  - `ID`: User registration number
  - `SESS`: Session ID
  - `DOM`: Domain identifier
  - `REF`: Referrer URL
  - `RD`: Redirect destination
- **Functionality**: 
  - Validates user credentials
  - Creates session records
  - Handles domain-specific redirects
  - Manages mobile vs desktop routing

#### WsLoginT.ashx
- **Purpose**: Alternative login service with additional features
- **Functionality**: Similar to WsLogin with enhanced session management

#### WsCLogin.ashx
- **Purpose**: Course-specific login service
- **Functionality**: Handles authentication for course access

#### bcplogout.ashx
- **Purpose**: Logout service
- **Functionality**: Cleans up sessions and redirects users

### Course Management

#### WsGetClass.ashx
- **Purpose**: Retrieves course information and prepares for launch
- **Parameters**:
  - `RID`: Registration ID
  - `UID`: User ID
  - `SES`: Session ID
  - `LANG`: Language code
- **Functionality**:
  - Validates course access
  - Checks KBA requirements
  - Generates course launch URLs
  - Manages SCORM vs HCI player routing
  - Handles document attachments

#### WsAcceptClass.ashx
- **Purpose**: Accepts and processes course completion
- **Functionality**:
  - Updates course status
  - Handles SCORM completion data
  - Manages progress tracking

#### WsLeaveClass.ashx
- **Purpose**: Handles course exit and cleanup
- **Functionality**:
  - Records exit events
  - Updates progress data
  - Manages session cleanup

#### WsGetClassAccess.ashx
- **Purpose**: Manages course access permissions
- **Functionality**:
  - Validates access rights
  - Handles KBA requirements
  - Manages course prerequisites

### Assessment Management

#### WsGetAssessment.ashx
- **Purpose**: Retrieves assessment information and prepares for launch
- **Functionality**:
  - Validates assessment access
  - Manages test sessions
  - Handles SCORM assessment data

#### WsLeaveAssessment.ashx
- **Purpose**: Handles assessment completion and cleanup
- **Functionality**:
  - Records assessment results
  - Updates completion status
  - Manages session cleanup

#### WsSaveKBA.ashx
- **Purpose**: Saves Knowledge-Based Assessment answers
- **Functionality**:
  - Stores KBA responses
  - Sends confirmation emails
  - Manages assessment completion

### Document Management

#### WsGetDocument.ashx
- **Purpose**: Retrieves document information and access tokens
- **Functionality**:
  - Validates document access
  - Generates secure access tokens
  - Manages document permissions

#### WsGetReport.ashx
- **Purpose**: Generates and retrieves reports
- **Functionality**:
  - Creates user reports
  - Manages report access
  - Handles report formatting

### Utility Services

#### NextAvail.ashx
- **Purpose**: Finds next available course sessions
- **Functionality**:
  - Searches available courses
  - Manages scheduling
  - Handles capacity management

#### etips.ashx
- **Purpose**: E-TIPS specific functionality
- **Functionality**:
  - Handles E-TIPS specific operations
  - Manages domain-specific features

#### RemoteLogin.ashx
- **Purpose**: Remote login functionality
- **Functionality**:
  - Handles external authentication
  - Manages cross-domain sessions

## Frontend Components

### HTML Interfaces

#### OpenClass.html
- Course launch interface
- Handles SCORM player integration
- Manages course documents and KBA questions

#### ClassAccess.html
- Course access management
- Handles registration and enrollment
- Manages course prerequisites

#### OpenAssessment.html
- Assessment launch interface
- Handles test delivery
- Manages assessment completion

#### OpenDocument.html
- Document viewing interface
- Handles secure document access
- Manages document permissions

#### FinishClass.html
- Course completion interface
- Handles completion confirmation
- Manages next steps

#### FinishAssessment.html
- Assessment completion interface
- Handles result display
- Manages certificate generation

#### OpenCertificate.html
- Certificate viewing interface
- Handles certificate display
- Manages certificate validation

#### OpenSurvey.html
- Survey interface
- Handles survey completion
- Manages survey data collection

#### AssessmentThanks.html
- Assessment completion confirmation
- Handles thank you messaging
- Manages user feedback

#### PlayerError.html
- Error handling interface
- Displays error messages
- Manages error recovery

#### logout.html
- Logout confirmation interface
- Handles logout process
- Manages session cleanup

## Configuration

### Web.config Settings

The system uses extensive configuration in Web.config:

#### Connection Strings
- `hcidb`: Main application database
- `siebeldb`: Siebel database
- `dms`: Document management system
- `reports`: Reporting database
- `email`: Email service database

#### App Settings
- `basepath`: Base file system path
- `dbuser`/`dbpass`: Database credentials
- `attachments`: Attachment storage path
- `minio-*`: MinIO object storage configuration
- Debug flags for each service
- `Lockoutcount`: Access control limits
- `LaunchProtocol`: Protocol for course launches

#### Logging Configuration
- Comprehensive log4net configuration
- Individual loggers for each service
- File and remote syslog appenders
- Debug logging for troubleshooting

### MinIO Integration

The system integrates with MinIO for object storage:

- **Configuration**: MinIO credentials and bucket settings in Web.config
- **Usage**: SCORM packages and course materials stored in MinIO
- **Access**: Secure URLs generated for content access
- **Integration**: AWS S3 SDK used for MinIO compatibility

## Security Features

### Authentication
- Session-based authentication
- Cookie management
- Cross-domain session handling
- Secure token generation

### Access Control
- User permission validation
- Course access restrictions
- Document access tokens
- Session timeout management

### Data Protection
- SQL injection prevention
- Input validation and sanitization
- Secure parameter handling
- Error message sanitization

## SCORM Integration

### Supported Versions
- SCORM 1.2
- SCORM 2004
- Custom HCI Player

### Features
- Launch URL generation
- Progress tracking
- Completion status management
- Score reporting
- Bookmark management

### Player Types
- **SCORM Player**: Standard SCORM-compliant player
- **HCI Player**: Custom player for specific content types
- **HTML5 Player**: Modern web-based player

## KBA (Knowledge-Based Assessment)

### Features
- Jurisdiction-specific questions
- Multi-language support
- Random question selection
- Answer tracking
- Email confirmations

### Integration
- Pre-course assessments
- Post-course evaluations
- Compliance tracking
- Certificate requirements

## Deployment

### Prerequisites
- Windows Server with IIS
- SQL Server 2012 or later
- .NET Framework 4.5.2
- MinIO server for object storage

### Installation Steps
1. Deploy web services to IIS
2. Run database schema script
3. Configure connection strings
4. Set up MinIO integration
5. Configure logging paths
6. Test all services

### Configuration
1. Update Web.config with environment-specific settings
2. Configure MinIO credentials and buckets
3. Set up database connections
4. Configure logging paths
5. Test service endpoints

## Monitoring and Logging

### Logging Levels
- **Event Log**: System events and errors
- **Debug Log**: Detailed debugging information
- **Performance Log**: Performance metrics

### Log Files
- Individual log files for each service
- Rolling file appenders
- Remote syslog integration
- Performance data collection

### Monitoring
- Service health checks
- Performance metrics
- Error tracking
- User activity logging

## API Documentation

### Common Parameters
- `UID`: User ID (registration number)
- `SES`: Session ID
- `RID`: Registration ID
- `LANG`: Language code (ENU, ESN)
- `DOMAIN`: Domain identifier
- `callback`: JSONP callback function

### Response Format
- JSON responses for most services
- Error handling with descriptive messages
- Status codes and result indicators
- Debug information when enabled

### Error Handling
- Comprehensive error logging
- User-friendly error messages
- Graceful degradation
- Recovery mechanisms

## Troubleshooting

### Common Issues
1. **Database Connection Errors**: Check connection strings and database availability
2. **MinIO Access Issues**: Verify credentials and bucket permissions
3. **Session Problems**: Check cookie settings and session management
4. **SCORM Launch Issues**: Verify player configuration and content URLs

### Debug Mode
- Enable debug logging in Web.config
- Check individual service log files
- Monitor performance metrics
- Review error logs

### Support
- Check log files for detailed error information
- Verify configuration settings
- Test individual service endpoints
- Review database connectivity

## Version History

- **v1.0**: Initial release with basic functionality
- **v1.1**: Added MinIO integration
- **v1.2**: Enhanced SCORM support
- **v1.3**: Improved KBA functionality
- **Current**: v1.4 with enhanced security and logging

## License

This software is proprietary and confidential. All rights reserved.

## Support

For technical support and questions, please contact the development team or refer to the internal documentation.
