using Microsoft.AspNetCore.Mvc;
using TeamTracker.API.Repositories;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly ExcelRepository _repo;
        private readonly IConfiguration _config;

        public AuthController(
            ExcelRepository repo,
            IConfiguration config)
        {
            _repo = repo;
            _config = config;
        }

        [HttpPost("login")]
        public async Task<IActionResult> Login(
            [FromBody] LoginRequest request)
        {
            Console.WriteLine("Login attempt:" + request.Username);
            var user = await _repo
                .GetUserByUsername(request.Username ?? "");
            Console.WriteLine("User Found" + (user != null ? "Yes" : "No"));
            if (user != null) 
            {
                Console.WriteLine("DB Password:" + user.PasswordHash);
                Console.WriteLine("Input Password:" + request.Password);
                Console.WriteLine("IsActive:" + user.IsActive);
            }

            if (user == null)
                return Unauthorized(new
                {
                    message = "Invalid credentials"
                });

            if (user.PasswordHash != request.Password)
                return Unauthorized(new
                {
                    message = "Invalid credentials"
                });

            if (!user.IsActive)
                return Unauthorized(new
                {
                    message = "User is inactive"
                });

            var token = GenerateToken(
                user.UserName, user.Role);

            return Ok(new
            {
                token = token,
                role = user.Role,
                username = user.UserName
            });
        }

        private string GenerateToken(
            string? username, string? role)
        {
            var key = new SymmetricSecurityKey(
                Encoding.UTF8.GetBytes(
                    _config["Jwt:Key"] ?? ""));

            var credentials =
                new SigningCredentials(key,
                SecurityAlgorithms.HmacSha256);

            var claims = new[]
            {
                new Claim(ClaimTypes.Name,
                    username ?? ""),
                new Claim(ClaimTypes.Role,
                    role ?? "")
            };

            var token = new JwtSecurityToken(
                issuer: _config["Jwt:Issuer"],
                audience: _config["Jwt:Audience"],
                claims: claims,
                expires: DateTime.Now.AddHours(8),
                signingCredentials: credentials
            );

            return new JwtSecurityTokenHandler()
                .WriteToken(token);
        }
    }

    public class LoginRequest
    {
        public string? Username { get; set; }
        public string? Password { get; set; }
    }
}
