using ADUserInfoService.Models;
using ADUserInfoService.Services;
using System.Runtime.Versioning;

namespace ADUserInfoService.Example
{
    [SupportedOSPlatform("windows")]
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("=== Active Directory User Information Service Example ===\n");

            try
            {
                await RunExamples();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        static async Task RunExamples()
        {
            Console.WriteLine("Choose connection method:");
            Console.WriteLine("1. Default domain (current user context)");
            Console.WriteLine("2. Specific domain");
            Console.WriteLine("3. Specific domain with credentials");
            Console.WriteLine("4. Specific domain with container/OU");
            Console.WriteLine("5. Full configuration (domain, container, credentials)");
            Console.Write("\nEnter choice (1-5): ");

            var choice = Console.ReadLine();

            using var adService = CreateADService(choice);

            if (adService == null)
            {
                Console.WriteLine("Invalid choice or error creating service.");
                return;
            }

            await DemonstrateFeatures(adService);
        }

        static ADUserService? CreateADService(string? choice)
        {
            switch (choice)
            {
                case "1":
                    Console.WriteLine("\nUsing default domain configuration...");
                    return new ADUserService();

                case "2":
                    Console.Write("Enter domain name: ");
                    var domain = Console.ReadLine();
                    if (string.IsNullOrEmpty(domain))
                        return null;
                    return new ADUserService(domain);

                case "3":
                    Console.Write("Enter domain name: ");
                    domain = Console.ReadLine();
                    Console.Write("Enter username: ");
                    var username = Console.ReadLine();
                    Console.Write("Enter password: ");
                    var password = ReadPassword();
                    if (string.IsNullOrEmpty(domain) || string.IsNullOrEmpty(username))
                        return null;
                    return new ADUserService(domain, username, password);

                case "4":
                    Console.Write("Enter domain name: ");
                    domain = Console.ReadLine();
                    Console.Write("Enter container/OU (e.g., OU=Users,DC=domain,DC=com): ");
                    var container = Console.ReadLine();
                    if (string.IsNullOrEmpty(domain))
                        return null;
                    return new ADUserService(domain, container ?? string.Empty);

                case "5":
                    Console.Write("Enter domain name: ");
                    domain = Console.ReadLine();
                    Console.Write("Enter container/OU: ");
                    container = Console.ReadLine();
                    Console.Write("Enter username: ");
                    username = Console.ReadLine();
                    Console.Write("Enter password: ");
                    password = ReadPassword();
                    if (string.IsNullOrEmpty(domain) || string.IsNullOrEmpty(username))
                        return null;
                    return new ADUserService(domain, container ?? string.Empty, username, password);

                default:
                    return null;
            }
        }

        static async Task DemonstrateFeatures(ADUserService adService)
        {
            while (true)
            {
                Console.WriteLine("\n=== Available Operations ===");
                Console.WriteLine("1. Get user by username");
                Console.WriteLine("2. Get user by email");
                Console.WriteLine("3. Get user by employee ID");
                Console.WriteLine("4. Check if user exists");
                Console.WriteLine("5. Check if user is enabled");
                Console.WriteLine("6. Check if user is locked out");
                Console.WriteLine("7. Get user groups");
                Console.WriteLine("8. Check if user is in group");
                Console.WriteLine("9. Search users");
                Console.WriteLine("10. Get users in group");
                Console.WriteLine("11. Get users by department");
                Console.WriteLine("12. Get direct reports");
                Console.WriteLine("13. Validate credentials");
                Console.WriteLine("14. Get user photo");
                Console.WriteLine("15. Export all users to Excel");
                Console.WriteLine("0. Exit");
                Console.Write("\nEnter choice: ");

                var operation = Console.ReadLine();

                if (operation == "0")
                    break;

                await ExecuteOperation(adService, operation);
            }
        }

        static async Task ExecuteOperation(ADUserService adService, string? operation)
        {
            try
            {
                switch (operation)
                {
                    case "1":
                        await GetUserByUsername(adService);
                        break;
                    case "2":
                        await GetUserByEmail(adService);
                        break;
                    case "3":
                        await GetUserByEmployeeId(adService);
                        break;
                    case "4":
                        await CheckUserExists(adService);
                        break;
                    case "5":
                        await CheckUserEnabled(adService);
                        break;
                    case "6":
                        await CheckUserLockedOut(adService);
                        break;
                    case "7":
                        await GetUserGroups(adService);
                        break;
                    case "8":
                        await CheckUserInGroup(adService);
                        break;
                    case "9":
                        await SearchUsers(adService);
                        break;
                    case "10":
                        await GetUsersInGroup(adService);
                        break;
                    case "11":
                        await GetUsersByDepartment(adService);
                        break;
                    case "12":
                        await GetDirectReports(adService);
                        break;
                    case "13":
                        await ValidateCredentials(adService);
                        break;
                    case "14":
                        await GetUserPhoto(adService);
                        break;
                    case "15":
                        await ExportAllUsersToExcel(adService);
                        break;
                    default:
                        Console.WriteLine("Invalid operation.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Operation failed: {ex.Message}");
            }
        }

        static async Task GetUserByUsername(ADUserService adService)
        {
            Console.Write("Enter username: ");
            var username = Console.ReadLine();
            if (string.IsNullOrEmpty(username))
                return;

            var user = await adService.GetUserByUsernameAsync(username);
            if (user != null)
            {
                DisplayUserInfo(user);
            }
            else
            {
                Console.WriteLine("User not found.");
            }
        }

        static async Task GetUserByEmail(ADUserService adService)
        {
            Console.Write("Enter email: ");
            var email = Console.ReadLine();
            if (string.IsNullOrEmpty(email))
                return;

            var users = await adService.GetUsersByEmailAsync(email);
            if (users != null && users.Count > 0)
            {
                Console.WriteLine($"\nFound {users.Count} user(s) with email '{email}':");
                Console.WriteLine(new string('-', 60));

                foreach (var user in users)
                {
                    DisplayUserInfo(user);
                    Console.WriteLine(new string('-', 60));
                }
            }
            else
            {
                Console.WriteLine("No users found with that email address.");
            }
        }

        static async Task GetUserByEmployeeId(ADUserService adService)
        {
            Console.Write("Enter employee ID: ");
            var employeeId = Console.ReadLine();
            if (string.IsNullOrEmpty(employeeId))
                return;

            var user = await adService.GetUserByEmployeeIdAsync(employeeId);
            if (user != null)
            {
                DisplayUserInfo(user);
            }
            else
            {
                Console.WriteLine("User not found.");
            }
        }

        static async Task CheckUserExists(ADUserService adService)
        {
            Console.Write("Enter username: ");
            var username = Console.ReadLine();
            if (string.IsNullOrEmpty(username))
                return;

            var exists = await adService.UserExistsAsync(username);
            Console.WriteLine($"User exists: {exists}");
        }

        static async Task CheckUserEnabled(ADUserService adService)
        {
            Console.Write("Enter username: ");
            var username = Console.ReadLine();
            if (string.IsNullOrEmpty(username))
                return;

            var enabled = await adService.IsUserEnabledAsync(username);
            Console.WriteLine($"User is enabled: {enabled}");
        }

        static async Task CheckUserLockedOut(ADUserService adService)
        {
            Console.Write("Enter username: ");
            var username = Console.ReadLine();
            if (string.IsNullOrEmpty(username))
                return;

            var lockedOut = await adService.IsUserLockedOutAsync(username);
            Console.WriteLine($"User is locked out: {lockedOut}");
        }

        static async Task GetUserGroups(ADUserService adService)
        {
            Console.Write("Enter username: ");
            var username = Console.ReadLine();
            if (string.IsNullOrEmpty(username))
                return;

            var groups = await adService.GetUserGroupsAsync(username);
            Console.WriteLine($"\nUser is member of {groups.Count} groups:");
            foreach (var group in groups)
            {
                Console.WriteLine($"  - {group}");
            }
        }

        static async Task CheckUserInGroup(ADUserService adService)
        {
            Console.Write("Enter username: ");
            var username = Console.ReadLine();
            Console.Write("Enter group name: ");
            var groupName = Console.ReadLine();
            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(groupName))
                return;

            var isMember = await adService.IsUserInGroupAsync(username, groupName);
            Console.WriteLine($"User is member of group: {isMember}");
        }

        static async Task SearchUsers(ADUserService adService)
        {
            Console.Write("Enter search term: ");
            var searchTerm = Console.ReadLine();
            Console.Write("Enter max results (default 100): ");
            var maxResultsStr = Console.ReadLine();
            if (string.IsNullOrEmpty(searchTerm))
                return;

            int maxResults = 100;
            if (!string.IsNullOrEmpty(maxResultsStr))
                int.TryParse(maxResultsStr, out maxResults);

            var users = await adService.SearchUsersAsync(searchTerm, maxResults);
            Console.WriteLine($"\nFound {users.Count} users:");
            foreach (var user in users)
            {
                Console.WriteLine($"  - {user.DisplayName} ({user.SamAccountName}) - {user.Email}");
            }
        }

        static async Task GetUsersInGroup(ADUserService adService)
        {
            Console.Write("Enter group name: ");
            var groupName = Console.ReadLine();
            if (string.IsNullOrEmpty(groupName))
                return;

            var users = await adService.GetUsersInGroupAsync(groupName);
            Console.WriteLine($"\nFound {users.Count} users in group:");
            foreach (var user in users)
            {
                Console.WriteLine($"  - {user.DisplayName} ({user.SamAccountName})");
            }
        }

        static async Task GetUsersByDepartment(ADUserService adService)
        {
            Console.Write("Enter department name: ");
            var department = Console.ReadLine();
            if (string.IsNullOrEmpty(department))
                return;

            var users = await adService.GetUsersByDepartmentAsync(department);
            Console.WriteLine($"\nFound {users.Count} users in department:");
            foreach (var user in users)
            {
                Console.WriteLine($"  - {user.DisplayName} ({user.SamAccountName}) - {user.Title}");
            }
        }

        static async Task GetDirectReports(ADUserService adService)
        {
            Console.Write("Enter manager username: ");
            var managerUsername = Console.ReadLine();
            if (string.IsNullOrEmpty(managerUsername))
                return;

            var reports = await adService.GetDirectReportsAsync(managerUsername);
            Console.WriteLine($"\nFound {reports.Count} direct reports:");
            foreach (var user in reports)
            {
                Console.WriteLine($"  - {user.DisplayName} ({user.SamAccountName}) - {user.Title}");
            }
        }

        static async Task ValidateCredentials(ADUserService adService)
        {
            Console.Write("Enter username: ");
            var username = Console.ReadLine();
            Console.Write("Enter password: ");
            var password = ReadPassword();
            if (string.IsNullOrEmpty(username))
                return;

            var valid = await adService.ValidateCredentialsAsync(username, password);
            Console.WriteLine($"\nCredentials are valid: {valid}");
        }

        static async Task GetUserPhoto(ADUserService adService)
        {
            Console.Write("Enter username: ");
            var username = Console.ReadLine();
            if (string.IsNullOrEmpty(username))
                return;

            var photo = await adService.GetUserPhotoAsync(username);
            if (photo != null)
            {
                Console.WriteLine($"User photo retrieved: {photo.Length} bytes");
                Console.Write("Save to file? (y/n): ");
                if (Console.ReadLine()?.ToLower() == "y")
                {
                    var fileName = $"{username}_photo.jpg";
                    await System.IO.File.WriteAllBytesAsync(fileName, photo);
                    Console.WriteLine($"Photo saved to {fileName}");
                }
            }
            else
            {
                Console.WriteLine("No photo found for user.");
            }
        }

        static void DisplayUserInfo(ADUserInfo user)
        {
            Console.WriteLine("\n=== User Information ===");
            Console.WriteLine($"Username: {user.SamAccountName}");
            Console.WriteLine($"Display Name: {user.DisplayName}");
            Console.WriteLine($"Email: {user.Email}");
            Console.WriteLine($"First Name: {user.FirstName}");
            Console.WriteLine($"Last Name: {user.LastName}");
            Console.WriteLine($"Employee ID: {user.EmployeeId}");
            Console.WriteLine($"Employee Number: {user.EmployeeNumber}");
            Console.WriteLine($"Department: {user.Department}");
            Console.WriteLine($"Title: {user.Title}");
            Console.WriteLine($"Company: {user.Company}");
            Console.WriteLine($"Manager: {user.Manager}");
            Console.WriteLine($"Office: {user.OfficeLocation}");
            Console.WriteLine($"Phone: {user.TelephoneNumber}");
            Console.WriteLine($"Mobile: {user.MobilePhone}");
            Console.WriteLine($"Enabled: {user.IsEnabled}");
            Console.WriteLine($"Locked Out: {user.IsLockedOut}");
            Console.WriteLine($"Last Logon: {user.LastLogonDate}");
            Console.WriteLine($"Created: {user.WhenCreated}");
            Console.WriteLine($"Modified: {user.WhenChanged}");
            Console.WriteLine($"Groups: {user.Groups.Count}");

            if (user.Groups.Count > 0 && user.Groups.Count <= 10)
            {
                Console.WriteLine("Group Memberships:");
                foreach (var group in user.Groups)
                {
                    Console.WriteLine($"  - {group}");
                }
            }
            else if (user.Groups.Count > 10)
            {
                Console.WriteLine($"  (Showing first 10 of {user.Groups.Count} groups)");
                for (int i = 0; i < 10; i++)
                {
                    Console.WriteLine($"  - {user.Groups[i]}");
                }
            }
        }

        static async Task ExportAllUsersToExcel(ADUserService adService)
        {
            Console.WriteLine("\n=== Export All Users to Excel ===");
            Console.Write("Enter directory path to save Excel file (press Enter for D:\\Temp): ");
            var directoryPath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                directoryPath = @"D:\Temp";
            }

            Console.WriteLine($"\nExporting all AD users to Excel...");
            Console.WriteLine($"Target directory: {directoryPath}");
            Console.WriteLine("Please wait, this may take a few moments...");

            try
            {
                var exportedFilePath = await adService.ExportAllUsersToExcelAsync(directoryPath);

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\n✓ Export completed successfully!");
                Console.ResetColor();
                Console.WriteLine($"File saved to: {exportedFilePath}");

                Console.Write("\nWould you like to open the file location? (y/n): ");
                if (Console.ReadLine()?.ToLower() == "y")
                {
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{exportedFilePath}\"");
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n✗ Export failed: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"   Inner Exception: {ex.InnerException.Message}");
                }
                Console.ResetColor();
            }
        }

        static string ReadPassword()
        {
            string password = "";
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(true);

                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    password += key.KeyChar;
                    Console.Write("*");
                }
                else if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                {
                    password = password.Substring(0, password.Length - 1);
                    Console.Write("\b \b");
                }
            }
            while (key.Key != ConsoleKey.Enter);

            Console.WriteLine();
            return password;
        }
    }
}