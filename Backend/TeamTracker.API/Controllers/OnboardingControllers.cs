using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using TeamTracker.API.Repositories;
using TeamTracker.API.Models;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class OnboardingController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;

        public OnboardingController(
            TeamTrackerRepository repo)
        {
            _repo = repo;
        }

        [HttpGet("steps")]
        public async Task<IActionResult> GetSteps()
        {
            var steps = await _repo
                .GetOnboardingSteps();
            return Ok(steps);
        }

        [HttpGet("teamsteps/{teamName}")]
        public async Task<IActionResult>
            GetStepsByTeam(string teamName)
        {
            var steps = await _repo
                .GetOnboardingStepsByTeam(teamName);
            return Ok(steps);
        }

        [HttpPost("steps")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> AddStep(
            [FromBody] OnboardingStep step)
        {
            await _repo.AddOnboardingStep(step);
            return Ok(new { message = "Step added" });
        }

        [HttpPatch("steps/{id}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> UpdateStep(
            int id,
            [FromBody] OnboardingStep step)
        {
            step.StepID = id;
            await _repo.UpdateOnboardingStep(step);
            return Ok(new
            {
                message = "Step updated"
            });
        }

        [HttpDelete("steps/{id}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> DeleteStep(
            int id)
        {
            await _repo
                .DeleteOnboardingStep(id);
            return Ok(new
            {
                message = "Step deleted"
            });
        }

        [HttpGet("progress")]
        public async Task<IActionResult> GetProgress()
        {
            var progress = await _repo
                .GetOnboardingProgress();
            return Ok(progress);
        }

        [HttpPatch("progress")]
        public async Task<IActionResult> UpdateProgress(
            [FromBody] OnboardingProgress progress)
        {
            await _repo.AddOrUpdateProgresswithNames(progress);
            return Ok(new
            {
                message = "Progress updated"
            });
        }

        [HttpPatch("progress/{candidateId}/{stepId}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult>
            UpdateCandidateProgress(
            int candidateId, int stepId,
            [FromBody] OnboardingProgress progress)
        {
            progress.CandidateID = candidateId;
            progress.StepID = stepId;
            await _repo.AddOrUpdateProgresswithNames(progress);
            return Ok(new
            {
                message = "Progress updated"
            });
        }

        [HttpDelete("progress/{candidateId}/{stepId}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult>
            DeleteCandidateProgress(
            int candidateId, int stepId)
        {
            await _repo.DeleteProgresswithNames(
                candidateId, stepId);
            return Ok(new
            {
                message = "Progress deleted"
            });
        }

        [HttpGet("statuswithnames")]
        public async Task<IActionResult> GetStatusWithNames()
        {
            var data = await _repo
                .GetOnboardingStatusWithNames();
            return Ok(data);
        }
    }
}
