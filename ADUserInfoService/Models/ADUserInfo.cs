using System;
using System.Collections.Generic;

namespace ADUserInfoService.Models
{
    public class ADUserInfo
    {
        public string? SamAccountName { get; set; }
        public string? UserPrincipalName { get; set; }
        public string? DisplayName { get; set; }
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public string? MiddleName { get; set; }
        public string? Email { get; set; }
        public string? EmployeeId { get; set; }
        public string? EmployeeNumber { get; set; }
        public string? EmployeeType { get; set; }
        
        public string? Department { get; set; }
        public string? Title { get; set; }
        public string? JobTitle { get; set; }
        public string? Company { get; set; }
        public string? Division { get; set; }
        public string? Organization { get; set; }
        
        public string? Manager { get; set; }
        public string? ManagerDistinguishedName { get; set; }
        public List<string> DirectReports { get; set; } = new List<string>();
        
        public string? OfficeLocation { get; set; }
        public string? StreetAddress { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public string? PostalCode { get; set; }
        public string? Country { get; set; }
        public string? CountryCode { get; set; }
        
        public string? TelephoneNumber { get; set; }
        public string? MobilePhone { get; set; }
        public string? HomePhone { get; set; }
        public string? FaxNumber { get; set; }
        public string? IpPhone { get; set; }
        public string? Pager { get; set; }
        public List<string> OtherTelephones { get; set; } = new List<string>();
        
        public string? HomeDirectory { get; set; }
        public string? HomeDrive { get; set; }
        public string? ProfilePath { get; set; }
        public string? ScriptPath { get; set; }
        
        public bool IsEnabled { get; set; }
        public bool IsLockedOut { get; set; }
        public bool PasswordNeverExpires { get; set; }
        public bool PasswordCannotChange { get; set; }
        public bool MustChangePasswordNextLogon { get; set; }
        public DateTime? AccountExpirationDate { get; set; }
        public DateTime? LastLogonDate { get; set; }
        public DateTime? LastPasswordSet { get; set; }
        public DateTime? BadPasswordTime { get; set; }
        public int BadPasswordCount { get; set; }
        public DateTime? LockoutTime { get; set; }
        
        public DateTime? WhenCreated { get; set; }
        public DateTime? WhenChanged { get; set; }
        
        public string? DistinguishedName { get; set; }
        public string? ObjectGuid { get; set; }
        public string? ObjectSid { get; set; }
        public string? Description { get; set; }
        
        public List<string> MemberOf { get; set; } = new List<string>();
        public List<string> Groups { get; set; } = new List<string>();
        
        public string? ThumbnailPhoto { get; set; }
        
        public string? ProxyAddresses { get; set; }
        public List<string> ProxyAddressList { get; set; } = new List<string>();
        
        public string? Info { get; set; }
        public string? PhysicalDeliveryOfficeName { get; set; }
        public string? PostOfficeBox { get; set; }
        
        public Dictionary<string, object> ExtensionAttributes { get; set; } = new Dictionary<string, object>();
        
        public Dictionary<string, object> AdditionalProperties { get; set; } = new Dictionary<string, object>();
        
        public override string ToString()
        {
            return $"{DisplayName} ({SamAccountName}) - {Email}";
        }
    }
}