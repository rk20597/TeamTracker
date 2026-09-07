using HROnboarding.API.Models;
using HROnboarding.API.Repositories;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace HROnboarding.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class TeamMemberController : ControllerBase
    {
        private readonly TeamTrackerRepository _repo;
        public TeamMemberController(TeamTrackerRepository repo)
        {
            _repo = repo;
        }

        [HttpGet("active")]
        [Authorize]
        public async Task<IActionResult> GetActive()
        {
            var members = await _repo.GetActiveMembers();
            
            return Ok(members);
        }

        [HttpGet("inactive")]
        [Authorize]
        public async Task<IActionResult> GetInactive()
        {
            var members = await _repo.GetInactiveMembers();

            return Ok(members);
        }

        [HttpGet("all")]
        [Authorize]
        public async Task<IActionResult> GetAll()
        {
            var members = await _repo.GetAllMembers();

            return Ok(members);
        }

        [HttpDelete("{srNo}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Delete(int srNo)
        {
            await _repo.DeleteTeamMember(srNo);
            return Ok(new { message = "Member deleted" });
        }

        [HttpPatch("{srNo}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Update(
    int srNo, [FromBody] TeamMember member)
        {
            member.SrNo = srNo;
            await _repo.UpdateTeamMember(member);
            return Ok(new { message = "Member updated" });
        }

        [HttpPost]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Add([FromBody] TeamMember member)
        { 
            await _repo.AddTeamMember(member);
            return Ok(new { message = "Member Added"});
        }



        //[HttpGet("debug")]
        //[AllowAnonymous]
        //public async Task<IActionResult> Debug()
        //{
        //    try
        //    {
        //        var all = await _repo.GetAllMembers();
        //        return Ok(new
        //        {
        //            totalCount = all.Count,
        //            activeCount = all.Count(m =>
        //                m.Status?.ToLower() == "active"),
        //            firstMember = all.FirstOrDefault()
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        return Ok(new { error = ex.Message });
        //    }
        //}
    }
}
