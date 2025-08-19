using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ADUserInfoService.Models;

namespace ADUserInfoService.Interfaces
{
    public interface IADUserService : IDisposable
    {
        ADUserInfo? GetUserByUsername(string username);
        
        Task<ADUserInfo?> GetUserByUsernameAsync(string username);
        
        List<ADUserInfo> GetUsersByEmail(string email);
        
        Task<List<ADUserInfo>> GetUsersByEmailAsync(string email);
        
        ADUserInfo? GetUserByEmail(string email);
        
        Task<ADUserInfo?> GetUserByEmailAsync(string email);
        
        ADUserInfo? GetUserByEmployeeId(string employeeId);
        
        Task<ADUserInfo?> GetUserByEmployeeIdAsync(string employeeId);
        
        ADUserInfo? GetUserByDistinguishedName(string distinguishedName);
        
        Task<ADUserInfo?> GetUserByDistinguishedNameAsync(string distinguishedName);
        
        bool UserExists(string username);
        
        Task<bool> UserExistsAsync(string username);
        
        bool UserExistsByEmail(string email);
        
        Task<bool> UserExistsByEmailAsync(string email);
        
        bool IsUserEnabled(string username);
        
        Task<bool> IsUserEnabledAsync(string username);
        
        bool IsUserLockedOut(string username);
        
        Task<bool> IsUserLockedOutAsync(string username);
        
        List<ADUserInfo> GetUsersInGroup(string groupName);
        
        Task<List<ADUserInfo>> GetUsersInGroupAsync(string groupName);
        
        List<ADUserInfo> GetUsersByDepartment(string department);
        
        Task<List<ADUserInfo>> GetUsersByDepartmentAsync(string department);
        
        List<ADUserInfo> GetDirectReports(string managerUsername);
        
        Task<List<ADUserInfo>> GetDirectReportsAsync(string managerUsername);
        
        List<string> GetUserGroups(string username);
        
        Task<List<string>> GetUserGroupsAsync(string username);
        
        bool IsUserInGroup(string username, string groupName);
        
        Task<bool> IsUserInGroupAsync(string username, string groupName);
        
        List<ADUserInfo> SearchUsers(string searchTerm, int maxResults = 100);
        
        Task<List<ADUserInfo>> SearchUsersAsync(string searchTerm, int maxResults = 100);
        
        bool ValidateCredentials(string username, string password);
        
        Task<bool> ValidateCredentialsAsync(string username, string password);
        
        byte[]? GetUserPhoto(string username);
        
        Task<byte[]?> GetUserPhotoAsync(string username);
        
        Task<string> ExportAllUsersToExcelAsync(string directoryPath);
    }
}