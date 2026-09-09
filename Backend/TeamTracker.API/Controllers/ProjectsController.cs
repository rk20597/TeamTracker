using TeamTracker.API.Repositories;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Authorization.Infrastructure;
using Microsoft.AspNetCore.Mvc;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(Roles = "Admin")]
    public class ProjectsController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;
        public ProjectsController(TeamTrackerRepository repo)
        {
            _repo = repo;
        }

        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            var projects = await _repo.GetProjects();
            return Ok(projects);
        }

        [HttpDelete("{projectName}")]
        public async Task<IActionResult> Delete(string projectName)
        {
            await _repo.DeleteProject(projectName);
            return Ok(new { message = "Project deleted" });
        }

    }
}
