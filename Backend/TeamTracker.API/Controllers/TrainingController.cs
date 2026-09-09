using TeamTracker.API.Repositories;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace TeamTracker.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class TrainingController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;
        public TrainingController(TeamTrackerRepository repo)
        {
            _repo = repo;
        }

        [HttpGet("courses/{domain}")]
        public async Task<IActionResult> GetCourses(string domain)
        {
            var validDomains = new[] { "Copilot", "AKS", "ROVO", "DataDog" };
            if (!validDomains.Contains(domain))
            {
                return BadRequest("Invalid Domain");
            }

            var courses = await _repo.GetTrainingCourses(domain);
            return Ok(courses);
        }

        [HttpGet("links/{domains}")]
        public async Task<IActionResult> GetLinks(string domains)
        {
            var validDomains = new[] { ".Net", "CRM", "iOS", "ETL", "RESRE", "ComplianceTrainingLinks" };
            if (!validDomains.Contains(domains))
                return BadRequest("Invalid Domain");

            var links = await _repo.GetTrainingLinks(domains);
            return Ok(links);
        }

        [HttpGet("mandatory")]
        public async Task<IActionResult> GetMandatory()
        { 
            var trainings = await _repo.GetMandatoryTrainings();
            return Ok(trainings);
        }

        [HttpDelete("links/{sheetName}/{topic}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> DeleteLink(
    string sheetName, string topic)
        {
            await _repo.DeleteTrainingLink(
                sheetName, topic);
            return Ok(new { message = "Link deleted" });
        }

        [HttpDelete("courses/{sheetName}/{title}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> DeleteCourse(
            string sheetName, string title)
        {
            await _repo.DeleteTrainingCourse(
                sheetName, title);
            return Ok(new { message = "Course deleted" });
        }


    }
}
