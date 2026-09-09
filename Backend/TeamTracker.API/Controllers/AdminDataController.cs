using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using TeamTracker.API.Repositories;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[Controller]")]
    [Authorize(Roles = "Admin")]
    public class AdminDataController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;
        public AdminDataController(TeamTrackerRepository repo) 
        { 
            _repo = repo;
        }

        [HttpGet("training-count")]
        public async Task<IActionResult> GetTrainingCount()
        {
            var data = await _repo.GetRawSheetData("Training Count");
            return Ok(data);
        }

        [HttpGet("team-demography")]
        public async Task<IActionResult> GetTeamDemography()
        {
            var data = await _repo.GetRawSheetData("Full Team Demography");
            return Ok(data);
        }

        [HttpGet("active-inactive")]
        public async Task<IActionResult> GetActiveInactive()
        {
            var data = await _repo.GetRawSheetData("Active-Inactive TMs");
            return Ok(data);
        }
    }
}
