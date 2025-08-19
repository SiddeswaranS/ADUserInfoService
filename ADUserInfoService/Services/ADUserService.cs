using ADUserInfoService.Interfaces;
using ADUserInfoService.Models;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Runtime.Versioning;

namespace ADUserInfoService.Services
{
    [SupportedOSPlatform("windows")]
    public class ADUserService : IADUserService
    {
        private readonly PrincipalContext _principalContext;
        private readonly string? _domain;
        private readonly string? _container;
        private readonly string? _username;
        private readonly string? _password;
        private bool _disposed = false;

        public ADUserService()
        {
            _principalContext = new PrincipalContext(ContextType.Domain);
        }

        public ADUserService(string domain)
        {
            _domain = domain;
            _principalContext = new PrincipalContext(ContextType.Domain, domain);
        }

        public ADUserService(string domain, string container)
        {
            _domain = domain;
            _container = container;
            _principalContext = new PrincipalContext(ContextType.Domain, domain, container);
        }

        public ADUserService(string domain, string username, string password)
        {
            _domain = domain;
            _username = username;
            _password = password;
            _principalContext = new PrincipalContext(ContextType.Domain, domain, username, password);
        }

        public ADUserService(string domain, string container, string username, string password)
        {
            _domain = domain;
            _container = container;
            _username = username;
            _password = password;
            _principalContext = new PrincipalContext(ContextType.Domain, domain, container, username, password);
        }

        public ADUserInfo? GetUserByUsername(string username)
        {
            try
            {
                using var userPrincipal = UserPrincipal.FindByIdentity(_principalContext, IdentityType.SamAccountName, username);
                if (userPrincipal == null)
                    return null;

                return GetUserInfo(userPrincipal);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error retrieving user by username '{username}'", ex);
            }
        }

        public async Task<ADUserInfo?> GetUserByUsernameAsync(string username)
        {
            return await Task.Run(() => GetUserByUsername(username));
        }

        public ADUserInfo? GetUserByEmail(string email)
        {
            try
            {
                using var userPrincipal = UserPrincipal.FindByIdentity(_principalContext, IdentityType.UserPrincipalName, email);
                if (userPrincipal == null)
                {
                    using var searcher = new PrincipalSearcher();
                    var user = new UserPrincipal(_principalContext) { EmailAddress = email };
                    searcher.QueryFilter = user;
                    var result = searcher.FindOne() as UserPrincipal;
                    if (result == null)
                        return null;

                    var info = GetUserInfo(result);
                    result.Dispose();
                    return info;
                }

                return GetUserInfo(userPrincipal);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error retrieving user by email '{email}'", ex);
            }
        }

        public async Task<ADUserInfo?> GetUserByEmailAsync(string email)
        {
            return await Task.Run(() => GetUserByEmail(email));
        }

        public ADUserInfo? GetUserByEmployeeId(string employeeId)
        {
            try
            {
                using var searcher = new PrincipalSearcher();
                var user = new UserPrincipal(_principalContext) { EmployeeId = employeeId };
                searcher.QueryFilter = user;
                var result = searcher.FindOne() as UserPrincipal;

                if (result == null)
                    return null;

                var info = GetUserInfo(result);
                result.Dispose();
                return info;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error retrieving user by employee ID '{employeeId}'", ex);
            }
        }

        public async Task<ADUserInfo?> GetUserByEmployeeIdAsync(string employeeId)
        {
            return await Task.Run(() => GetUserByEmployeeId(employeeId));
        }

        public ADUserInfo? GetUserByDistinguishedName(string distinguishedName)
        {
            try
            {
                using var userPrincipal = UserPrincipal.FindByIdentity(_principalContext, IdentityType.DistinguishedName, distinguishedName);
                if (userPrincipal == null)
                    return null;

                return GetUserInfo(userPrincipal);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error retrieving user by distinguished name '{distinguishedName}'", ex);
            }
        }

        public async Task<ADUserInfo?> GetUserByDistinguishedNameAsync(string distinguishedName)
        {
            return await Task.Run(() => GetUserByDistinguishedName(distinguishedName));
        }

        public bool UserExists(string username)
        {
            try
            {
                using var user = UserPrincipal.FindByIdentity(_principalContext, IdentityType.SamAccountName, username);
                return user != null;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> UserExistsAsync(string username)
        {
            return await Task.Run(() => UserExists(username));
        }

        public bool UserExistsByEmail(string email)
        {
            try
            {
                using var user = UserPrincipal.FindByIdentity(_principalContext, IdentityType.UserPrincipalName, email);
                if (user != null)
                    return true;

                using var searcher = new PrincipalSearcher();
                var userFilter = new UserPrincipal(_principalContext) { EmailAddress = email };
                searcher.QueryFilter = userFilter;
                var result = searcher.FindOne();
                return result != null;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> UserExistsByEmailAsync(string email)
        {
            return await Task.Run(() => UserExistsByEmail(email));
        }

        public bool IsUserEnabled(string username)
        {
            try
            {
                using var user = UserPrincipal.FindByIdentity(_principalContext, IdentityType.SamAccountName, username);
                return user?.Enabled ?? false;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> IsUserEnabledAsync(string username)
        {
            return await Task.Run(() => IsUserEnabled(username));
        }

        public bool IsUserLockedOut(string username)
        {
            try
            {
                using var user = UserPrincipal.FindByIdentity(_principalContext, IdentityType.SamAccountName, username);
                return user?.IsAccountLockedOut() ?? false;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> IsUserLockedOutAsync(string username)
        {
            return await Task.Run(() => IsUserLockedOut(username));
        }

        public List<ADUserInfo> GetUsersInGroup(string groupName)
        {
            var users = new List<ADUserInfo>();
            try
            {
                using var group = GroupPrincipal.FindByIdentity(_principalContext, groupName);
                if (group == null)
                    return users;

                foreach (var member in group.GetMembers(true))
                {
                    if (member is UserPrincipal userPrincipal)
                    {
                        users.Add(GetUserInfo(userPrincipal));
                        userPrincipal.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error retrieving users in group '{groupName}'", ex);
            }

            return users;
        }

        public async Task<List<ADUserInfo>> GetUsersInGroupAsync(string groupName)
        {
            return await Task.Run(() => GetUsersInGroup(groupName));
        }

        public List<ADUserInfo> GetUsersByDepartment(string department)
        {
            var users = new List<ADUserInfo>();
            try
            {
                var ldapPath = GetLdapPath();
                using var directoryEntry = new DirectoryEntry(ldapPath, _username, _password);
                using var searcher = new DirectorySearcher(directoryEntry)
                {
                    Filter = $"(&(objectClass=user)(department={department}))",
                    PageSize = 1000
                };

                searcher.PropertiesToLoad.Add("samAccountName");

                foreach (SearchResult result in searcher.FindAll())
                {
                    var samAccountName = result.Properties["samAccountName"][0]?.ToString();
                    if (!string.IsNullOrEmpty(samAccountName))
                    {
                        var userInfo = GetUserByUsername(samAccountName);
                        if (userInfo != null)
                            users.Add(userInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error retrieving users by department '{department}'", ex);
            }

            return users;
        }

        public async Task<List<ADUserInfo>> GetUsersByDepartmentAsync(string department)
        {
            return await Task.Run(() => GetUsersByDepartment(department));
        }

        public List<ADUserInfo> GetDirectReports(string managerUsername)
        {
            var directReports = new List<ADUserInfo>();
            try
            {
                var manager = GetUserByUsername(managerUsername);
                if (manager == null)
                    return directReports;

                foreach (var reportDn in manager.DirectReports)
                {
                    var report = GetUserByDistinguishedName(reportDn);
                    if (report != null)
                        directReports.Add(report);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error retrieving direct reports for '{managerUsername}'", ex);
            }

            return directReports;
        }

        public async Task<List<ADUserInfo>> GetDirectReportsAsync(string managerUsername)
        {
            return await Task.Run(() => GetDirectReports(managerUsername));
        }

        public List<string> GetUserGroups(string username)
        {
            var groups = new List<string>();
            try
            {
                using var user = UserPrincipal.FindByIdentity(_principalContext, IdentityType.SamAccountName, username);
                if (user == null)
                    return groups;

                var principalGroups = user.GetAuthorizationGroups();
                foreach (var group in principalGroups)
                {
                    groups.Add(group.Name);
                    group.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error retrieving groups for user '{username}'", ex);
            }

            return groups;
        }

        public async Task<List<string>> GetUserGroupsAsync(string username)
        {
            return await Task.Run(() => GetUserGroups(username));
        }

        public bool IsUserInGroup(string username, string groupName)
        {
            try
            {
                using var user = UserPrincipal.FindByIdentity(_principalContext, IdentityType.SamAccountName, username);
                if (user == null)
                    return false;

                using var group = GroupPrincipal.FindByIdentity(_principalContext, groupName);
                if (group == null)
                    return false;

                return user.IsMemberOf(group);
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> IsUserInGroupAsync(string username, string groupName)
        {
            return await Task.Run(() => IsUserInGroup(username, groupName));
        }

        public List<ADUserInfo> SearchUsers(string searchTerm, int maxResults = 100)
        {
            var users = new List<ADUserInfo>();
            try
            {
                using var searcher = new PrincipalSearcher();
                var userFilter = new UserPrincipal(_principalContext)
                {
                    Name = $"*{searchTerm}*"
                };
                searcher.QueryFilter = userFilter;

                var results = searcher.FindAll().Take(maxResults);
                foreach (UserPrincipal user in results)
                {
                    users.Add(GetUserInfo(user));
                    user.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error searching for users with term '{searchTerm}'", ex);
            }

            return users;
        }

        public async Task<List<ADUserInfo>> SearchUsersAsync(string searchTerm, int maxResults = 100)
        {
            return await Task.Run(() => SearchUsers(searchTerm, maxResults));
        }

        public bool ValidateCredentials(string username, string password)
        {
            try
            {
                return _principalContext.ValidateCredentials(username, password);
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> ValidateCredentialsAsync(string username, string password)
        {
            return await Task.Run(() => ValidateCredentials(username, password));
        }

        public byte[]? GetUserPhoto(string username)
        {
            try
            {
                using var user = UserPrincipal.FindByIdentity(_principalContext, IdentityType.SamAccountName, username);
                if (user == null)
                    return null;

                using var directoryEntry = user.GetUnderlyingObject() as DirectoryEntry;
                if (directoryEntry?.Properties.Contains("thumbnailPhoto") == true)
                {
                    return directoryEntry.Properties["thumbnailPhoto"].Value as byte[];
                }
            }
            catch
            {
                // Silently fail if photo cannot be retrieved
            }

            return null;
        }

        public async Task<byte[]?> GetUserPhotoAsync(string username)
        {
            return await Task.Run(() => GetUserPhoto(username));
        }

        private ADUserInfo GetUserInfo(UserPrincipal userPrincipal)
        {
            var userInfo = new ADUserInfo
            {
                SamAccountName = userPrincipal.SamAccountName,
                UserPrincipalName = userPrincipal.UserPrincipalName,
                DisplayName = userPrincipal.DisplayName,
                FirstName = userPrincipal.GivenName,
                LastName = userPrincipal.Surname,
                MiddleName = userPrincipal.MiddleName,
                Email = userPrincipal.EmailAddress,
                EmployeeId = userPrincipal.EmployeeId,
                IsEnabled = userPrincipal.Enabled ?? false,
                IsLockedOut = userPrincipal.IsAccountLockedOut(),
                AccountExpirationDate = userPrincipal.AccountExpirationDate,
                LastLogonDate = userPrincipal.LastLogon,
                LastPasswordSet = userPrincipal.LastPasswordSet,
                BadPasswordCount = userPrincipal.BadLogonCount,
                DistinguishedName = userPrincipal.DistinguishedName,
                ObjectGuid = userPrincipal.Guid?.ToString(),
                ObjectSid = userPrincipal.Sid?.ToString(),
                Description = userPrincipal.Description,
                TelephoneNumber = userPrincipal.VoiceTelephoneNumber,
                HomeDirectory = userPrincipal.HomeDirectory,
                HomeDrive = userPrincipal.HomeDrive
            };

            try
            {
                using var directoryEntry = userPrincipal.GetUnderlyingObject() as DirectoryEntry;
                if (directoryEntry != null)
                {
                    PopulateExtendedProperties(userInfo, directoryEntry);
                }
            }
            catch
            {
                // Silently fail for extended properties
            }

            try
            {
                var groups = userPrincipal.GetAuthorizationGroups();
                foreach (var group in groups)
                {
                    userInfo.Groups.Add(group.Name);
                    userInfo.MemberOf.Add(group.DistinguishedName);
                    group.Dispose();
                }
            }
            catch
            {
                // Silently fail for group membership
            }

            return userInfo;
        }

        private void PopulateExtendedProperties(ADUserInfo userInfo, DirectoryEntry directoryEntry)
        {
            userInfo.Department = GetPropertyValue(directoryEntry, "department");
            userInfo.Title = GetPropertyValue(directoryEntry, "title");
            userInfo.JobTitle = GetPropertyValue(directoryEntry, "title");
            userInfo.Company = GetPropertyValue(directoryEntry, "company");
            userInfo.Division = GetPropertyValue(directoryEntry, "division");
            userInfo.Organization = GetPropertyValue(directoryEntry, "o");
            userInfo.EmployeeNumber = GetPropertyValue(directoryEntry, "employeeNumber");
            userInfo.EmployeeType = GetPropertyValue(directoryEntry, "employeeType");

            userInfo.Manager = GetPropertyValue(directoryEntry, "manager");
            userInfo.ManagerDistinguishedName = userInfo.Manager;

            var directReports = directoryEntry.Properties["directReports"];
            if (directReports != null)
            {
                foreach (var report in directReports)
                {
                    userInfo.DirectReports.Add(report.ToString()!);
                }
            }

            userInfo.OfficeLocation = GetPropertyValue(directoryEntry, "physicalDeliveryOfficeName");
            userInfo.StreetAddress = GetPropertyValue(directoryEntry, "streetAddress");
            userInfo.City = GetPropertyValue(directoryEntry, "l");
            userInfo.State = GetPropertyValue(directoryEntry, "st");
            userInfo.PostalCode = GetPropertyValue(directoryEntry, "postalCode");
            userInfo.Country = GetPropertyValue(directoryEntry, "co");
            userInfo.CountryCode = GetPropertyValue(directoryEntry, "c");

            userInfo.MobilePhone = GetPropertyValue(directoryEntry, "mobile");
            userInfo.HomePhone = GetPropertyValue(directoryEntry, "homePhone");
            userInfo.FaxNumber = GetPropertyValue(directoryEntry, "facsimileTelephoneNumber");
            userInfo.IpPhone = GetPropertyValue(directoryEntry, "ipPhone");
            userInfo.Pager = GetPropertyValue(directoryEntry, "pager");

            var otherTelephones = directoryEntry.Properties["otherTelephone"];
            if (otherTelephones != null)
            {
                foreach (var phone in otherTelephones)
                {
                    userInfo.OtherTelephones.Add(phone.ToString()!);
                }
            }

            userInfo.ProfilePath = GetPropertyValue(directoryEntry, "profilePath");
            userInfo.ScriptPath = GetPropertyValue(directoryEntry, "scriptPath");

            var userAccountControl = GetPropertyValue(directoryEntry, "userAccountControl");
            if (!string.IsNullOrEmpty(userAccountControl) && int.TryParse(userAccountControl, out int uac))
            {
                userInfo.PasswordNeverExpires = (uac & 0x10000) != 0;
                userInfo.PasswordCannotChange = (uac & 0x40) != 0;
                userInfo.MustChangePasswordNextLogon = (uac & 0x800000) != 0;
            }

            userInfo.WhenCreated = GetDateTimeProperty(directoryEntry, "whenCreated");
            userInfo.WhenChanged = GetDateTimeProperty(directoryEntry, "whenChanged");
            userInfo.BadPasswordTime = GetDateTimeProperty(directoryEntry, "badPasswordTime");
            userInfo.LockoutTime = GetDateTimeProperty(directoryEntry, "lockoutTime");

            if (directoryEntry.Properties.Contains("thumbnailPhoto"))
            {
                var photo = directoryEntry.Properties["thumbnailPhoto"].Value as byte[];
                if (photo != null)
                {
                    userInfo.ThumbnailPhoto = Convert.ToBase64String(photo);
                }
            }

            userInfo.ProxyAddresses = GetPropertyValue(directoryEntry, "proxyAddresses");
            var proxyAddresses = directoryEntry.Properties["proxyAddresses"];
            if (proxyAddresses != null)
            {
                foreach (var address in proxyAddresses)
                {
                    userInfo.ProxyAddressList.Add(address.ToString()!);
                }
            }

            userInfo.Info = GetPropertyValue(directoryEntry, "info");
            userInfo.PhysicalDeliveryOfficeName = GetPropertyValue(directoryEntry, "physicalDeliveryOfficeName");
            userInfo.PostOfficeBox = GetPropertyValue(directoryEntry, "postOfficeBox");

            for (int i = 1; i <= 15; i++)
            {
                var extensionAttribute = GetPropertyValue(directoryEntry, $"extensionAttribute{i}");
                if (!string.IsNullOrEmpty(extensionAttribute))
                {
                    userInfo.ExtensionAttributes[$"extensionAttribute{i}"] = extensionAttribute;
                }
            }
        }

        private string? GetPropertyValue(DirectoryEntry entry, string propertyName)
        {
            if (entry.Properties.Contains(propertyName))
            {
                var value = entry.Properties[propertyName].Value;
                return value?.ToString();
            }
            return null;
        }

        private DateTime? GetDateTimeProperty(DirectoryEntry entry, string propertyName)
        {
            if (entry.Properties.Contains(propertyName))
            {
                var value = entry.Properties[propertyName].Value;
                if (value is DateTime dateTime)
                    return dateTime;

                if (value is long fileTime && fileTime > 0)
                {
                    try
                    {
                        return DateTime.FromFileTime(fileTime);
                    }
                    catch
                    {
                        // Invalid file time
                    }
                }
            }
            return null;
        }

        private string GetLdapPath()
        {
            if (!string.IsNullOrEmpty(_container))
                return $"LDAP://{_domain}/{_container}";
            else if (!string.IsNullOrEmpty(_domain))
                return $"LDAP://{_domain}";
            else
                return "LDAP://RootDSE";
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _principalContext?.Dispose();
                }
                _disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}