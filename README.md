# ADUserInfoService

A comprehensive .NET library and console application for interacting with Active Directory to retrieve and export user information.

## Features

- **User Information Retrieval**: Get detailed user information from Active Directory
- **Multiple Search Options**: Search by username, email, employee ID, department, etc.
- **Group Management**: Check group memberships and retrieve users in specific groups
- **User Status Checks**: Verify if users are enabled, locked out, or exist in AD
- **Credential Validation**: Validate user credentials against Active Directory
- **Excel Export**: Export all AD users to Excel format with timestamp
- **Async Support**: All operations support both synchronous and asynchronous execution
- **Flexible Connection Options**: Multiple ways to connect to AD with different authentication methods

## Prerequisites

- .NET 9.0 or later
- Windows operating system (required for System.DirectoryServices)
- Active Directory domain environment
- Appropriate permissions to read from Active Directory

## Installation

### NuGet Packages Required

```xml
<PackageReference Include="System.DirectoryServices.AccountManagement" Version="9.0.8" />
<PackageReference Include="EPPlus" Version="7.7.3" />
```

## Project Structure

```
ADUserInfoService/
├── ADUserInfoService/              # Core library project
│   ├── Interfaces/
│   │   └── IADUserService.cs       # Service interface definition
│   ├── Models/
│   │   └── ADUserInfo.cs           # User information model
│   ├── Services/
│   │   └── ADUserService.cs        # Main AD service implementation
│   └── ADUserInfoService.csproj
├── ADUserInfoService.Sample/       # Console application for testing
│   ├── Program.cs
│   └── ADUserInfoService.Sample.csproj
└── README.md
```

## Usage

### Basic Usage

```csharp
using ADUserInfoService.Services;
using ADUserInfoService.Models;

// Create service with default domain
using var adService = new ADUserService();

// Get user by username
var user = await adService.GetUserByUsernameAsync("john.doe");
if (user != null)
{
    Console.WriteLine($"Display Name: {user.DisplayName}");
    Console.WriteLine($"Email: {user.Email}");
    Console.WriteLine($"Department: {user.Department}");
}
```

### Connection Options

```csharp
// 1. Default domain (current user context)
var adService = new ADUserService();

// 2. Specific domain
var adService = new ADUserService("contoso.com");

// 3. Specific domain with credentials
var adService = new ADUserService("contoso.com", "username", "password");

// 4. Specific domain with container/OU
var adService = new ADUserService("contoso.com", "OU=Users,DC=contoso,DC=com");

// 5. Full configuration
var adService = new ADUserService("contoso.com", "OU=Users,DC=contoso,DC=com", "username", "password");
```

### Available Operations

#### User Retrieval
```csharp
// Get user by username
var user = await adService.GetUserByUsernameAsync("username");

// Get users by email
var users = await adService.GetUsersByEmailAsync("user@domain.com");

// Get user by employee ID
var user = await adService.GetUserByEmployeeIdAsync("EMP001");

// Search users
var users = await adService.SearchUsersAsync("john", maxResults: 100);
```

#### User Status Checks
```csharp
// Check if user exists
bool exists = await adService.UserExistsAsync("username");

// Check if user is enabled
bool isEnabled = await adService.IsUserEnabledAsync("username");

// Check if user is locked out
bool isLocked = await adService.IsUserLockedOutAsync("username");

// Validate credentials
bool isValid = await adService.ValidateCredentialsAsync("username", "password");
```

#### Group Operations
```csharp
// Get user groups
var groups = await adService.GetUserGroupsAsync("username");

// Check if user is in specific group
bool isMember = await adService.IsUserInGroupAsync("username", "GroupName");

// Get all users in a group
var users = await adService.GetUsersInGroupAsync("GroupName");
```

#### Department and Reporting
```csharp
// Get users by department
var users = await adService.GetUsersByDepartmentAsync("IT");

// Get direct reports for a manager
var reports = await adService.GetDirectReportsAsync("manager.username");
```

#### Export to Excel
```csharp
// Export all AD users to Excel with timestamp
string filePath = await adService.ExportAllUsersToExcelAsync(@"C:\Exports");
// Creates file like: C:\Exports\ADUsers_20250819_143025.xlsx
```

## ADUserInfo Model Properties

The `ADUserInfo` class contains the following properties:

### Basic Information
- `SamAccountName` - User login name
- `UserPrincipalName` - User principal name (UPN)
- `DisplayName` - Display name
- `FirstName`, `LastName`, `MiddleName` - Name components
- `Email` - Email address

### Employment Information
- `EmployeeId` - Employee ID
- `EmployeeNumber` - Employee number
- `EmployeeType` - Type of employee
- `Department` - Department name
- `Title` / `JobTitle` - Job title
- `Company` - Company name
- `Division` - Division
- `Organization` - Organization
- `Manager` - Manager's distinguished name

### Contact Information
- `TelephoneNumber` - Office phone
- `MobilePhone` - Mobile phone
- `HomePhone` - Home phone
- `FaxNumber` - Fax number
- `IpPhone` - IP phone
- `Pager` - Pager number
- `OtherTelephones` - List of other phone numbers

### Location Information
- `OfficeLocation` - Office location
- `StreetAddress` - Street address
- `City` - City
- `State` - State/Province
- `PostalCode` - Postal/ZIP code
- `Country` - Country
- `CountryCode` - Country code

### Account Status
- `IsEnabled` - Account enabled status
- `IsLockedOut` - Account locked status
- `PasswordNeverExpires` - Password expiry setting
- `PasswordCannotChange` - Password change permission
- `MustChangePasswordNextLogon` - Password change requirement
- `AccountExpirationDate` - Account expiration date
- `LastLogonDate` - Last logon timestamp
- `LastPasswordSet` - Password last set date
- `BadPasswordCount` - Failed login attempts
- `BadPasswordTime` - Last failed login time
- `LockoutTime` - Account lockout time

### System Information
- `DistinguishedName` - DN in Active Directory
- `ObjectGuid` - Unique object GUID
- `ObjectSid` - Security identifier
- `WhenCreated` - Creation timestamp
- `WhenChanged` - Last modification timestamp
- `Description` - User description

### Group Membership
- `MemberOf` - List of group DNs
- `Groups` - List of group names

### Additional Properties
- `HomeDirectory` - Home folder path
- `HomeDrive` - Home drive letter
- `ProfilePath` - Profile path
- `ScriptPath` - Logon script path
- `ThumbnailPhoto` - User photo (Base64)
- `ProxyAddresses` - Proxy addresses
- `ExtensionAttributes` - Extension attributes
- `AdditionalProperties` - Other custom properties

## Excel Export Format

The Excel export includes the following columns:
- SamAccountName
- UserPrincipalName
- DisplayName
- FirstName, LastName, MiddleName
- Email
- EmployeeId, EmployeeNumber, EmployeeType
- Department, Title, Company, Division, Organization
- Manager
- OfficeLocation, StreetAddress, City, State, PostalCode, Country
- TelephoneNumber, MobilePhone, HomePhone
- IsEnabled, IsLockedOut
- LastLogonDate, WhenCreated
- DistinguishedName

## Console Application Usage

Run the sample console application:

```bash
dotnet run --project ADUserInfoService.Sample
```

### Menu Options
1. Get user by username
2. Get user by email
3. Get user by employee ID
4. Check if user exists
5. Check if user is enabled
6. Check if user is locked out
7. Get user groups
8. Check if user is in group
9. Search users
10. Get users in group
11. Get users by department
12. Get direct reports
13. Validate credentials
14. Get user photo
15. **Export all users to Excel**

## Error Handling

The library includes comprehensive error handling:
- Specific exceptions for AD operations
- Detailed error messages
- Inner exception information
- Graceful handling of missing permissions

## Security Considerations

- Always use secure methods to handle passwords
- Store credentials securely (use Windows Credential Manager or Azure Key Vault)
- Ensure proper permissions are granted for AD operations
- Use service accounts with minimum required permissions
- Enable auditing for sensitive operations

## License

This project is licensed under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues, questions, or suggestions, please create an issue in the GitHub repository.

## Requirements

- Windows OS (required for System.DirectoryServices)
- .NET 9.0 Runtime
- Active Directory environment
- Network connectivity to domain controllers
- Read permissions in Active Directory

## Known Limitations

- Only works on Windows due to System.DirectoryServices dependency
- Requires domain-joined machine or proper network access to AD
- Large exports may take time depending on AD size
- Some properties may be null if not populated in AD

## Version History

- **1.0.0** - Initial release with basic AD operations
- **1.1.0** - Added Excel export functionality with EPPlus
- **1.2.0** - Added async support for all operations