using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using TeamTracker.API.Repositories;
using TeamTracker.API.Models;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class PRSDController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;

        public PRSDController(
            TeamTrackerRepository repo)
        {
            _repo = repo;
        }

        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            var data = await _repo.GetPRSD();
            return Ok(data);
        }

        [HttpPost]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Add(
            [FromBody] PRSD prsd)
        {
            await _repo.AddPRSD(prsd);
            return Ok(new { message = "PRSD added" });
        }

        [HttpPatch]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Update(
            [FromBody] PRSD prsd)
        {
            await _repo.UpdatePRSD(prsd);
            return Ok(new
            {
                message = "PRSD updated"
            });
        }

        [HttpDelete]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Delete(
            [FromQuery] string lob,
            [FromQuery] string trackName)
        {
            await _repo.DeletePRSD(lob, trackName);
            return Ok(new
            {
                message = "PRSD deleted"
            });
        }
    }
}

