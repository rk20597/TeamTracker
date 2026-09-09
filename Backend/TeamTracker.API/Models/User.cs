using Microsoft.AspNetCore.Identity;

namespace TeamTracker.API.Models
{
    public class User
    {
        public int UserID { get; set; }
        public string? UserName { get; set; }
        public string? PasswordHash { get; set; }

        public string? Role { get; set; }

        public bool IsActive { get; set; }
    }
}
