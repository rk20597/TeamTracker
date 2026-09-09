using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using TeamTracker.API.Repositories;
using TeamTracker.API.Models;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class OffboardedController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;

        public OffboardedController(
            TeamTrackerRepository repo)
        {
            _repo = repo;
        }

        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            var data = await _repo.GetOffboarded();
            return Ok(data);
        }

        [HttpPost]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Add(
            [FromBody] Offboarded member)
        {
            await _repo.AddOffboarded(member);
            return Ok(new
            {
                message = "Member added"
            });
        }

        [HttpDelete("{id}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Delete(
            int id)
        {
            await _repo.DeleteOffboarded(id);
            return Ok(new
            {
                message = "Member deleted"
            });
        }

        [HttpPatch("{id}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Update(
    int id, [FromBody] Offboarded member)
        {
            member.CandidateID = id;
            await _repo.UpdateOffboarded(member);
            return Ok(new { message = "Updated" });
        }

    }
}

