using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using TeamTracker.API.Repositories;
using TeamTracker.API.Models;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(Roles = "Admin")]
    public class UsersController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;

        public UsersController(
            TeamTrackerRepository repo)
        {
            _repo = repo;
        }

        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            var data = await _repo.GetAllUsers();
            return Ok(data);
        }

        [HttpPost]
        public async Task<IActionResult> Add(
            [FromBody] User user)
        {
            await _repo.AddUser(user);
            return Ok(new
            {
                message = "User added"
            });
        }

        [HttpPatch("{id}")]
        public async Task<IActionResult> Update(
            int id, [FromBody] User user)
        {
            user.UserID = id;
            await _repo.UpdateUser(user);
            return Ok(new
            {
                message = "User updated"
            });
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(
            int id)
        {
            await _repo.DeleteUser(id);
            return Ok(new
            {
                message = "User deleted"
            });
        }
    }
}

