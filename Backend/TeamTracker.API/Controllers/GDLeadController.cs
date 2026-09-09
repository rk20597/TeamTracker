using TeamTracker.API.Repositories;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class GDLeadController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;
        public GDLeadController(TeamTrackerRepository repo)
        {
            _repo = repo;
        }

        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            var data = await _repo.GetGDLeadData();
            return Ok(data);
        }

        [HttpDelete("{skillSet}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Delete(
    string skillSet)
        {
            await _repo.DeleteGDLead(skillSet);
            return Ok(new { message = "GD Lead deleted" });
        }


    }
}
