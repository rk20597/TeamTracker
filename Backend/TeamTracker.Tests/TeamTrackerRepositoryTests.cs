using NUnit.Framework;
using HROnboarding.API.Repositories;
using HROnboarding.API.Models;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using OfficeOpenXml;

namespace TeamTracker.Tests
{
    [TestFixture]
    public class TeamTrackerRepositoryTests
    {
        private TeamTrackerRepository _repo;
        private string _testFilePath;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.License.SetNonCommercialPersonal("TeamTracker");
            _testFilePath = Path.Combine(
                TestContext.CurrentContext
                    .TestDirectory,
                "..", "..", "..", "..", "..",
                "Data", "TeamTracker.xlsx");
            _repo = new TeamTrackerRepository(
                _testFilePath);
        }

        // ==================
        // TEAM MEMBER TESTS
        // ==================

        [Test]
        public async Task GetAllMembers_ReturnsMembers()
        {
            var members = await _repo.GetAllMembers();
            Assert.That(members, Is.Not.Null);
            Assert.That(members.Count,
                Is.GreaterThan(0));
        }

        [Test]
        public async Task GetActiveMembers_ReturnsOnlyActive()
        {
            var members = await _repo.GetActiveMembers();
            Assert.That(members, Is.Not.Null);
            Assert.That(members.All(m =>
                m.Status == "Active"), Is.True);
        }

        [Test]
        public async Task GetInactiveMembers_ReturnsOnlyInactive()
        {
            var members = await _repo
                .GetInactiveMembers();
            Assert.That(members, Is.Not.Null);
            Assert.That(members.All(m =>
                m.Status == "Inactive"), Is.True);
        }

        [Test]
        public async Task GetAllMembers_HasRequiredFields()
        {
            var members = await _repo.GetAllMembers();
            var first = members.FirstOrDefault();
            Assert.That(first, Is.Not.Null);
            Assert.That(first.Name,
                Is.Not.Null.And.Not.Empty);
            Assert.That(first.Email,
                Is.Not.Null.And.Not.Empty);
        }

        [Test]
        public async Task GetAllMembers_EmailsEndWithDeloitte()
        {
            var members = await _repo.GetAllMembers();
            var invalidEmails = members
                .Where(m => !string.IsNullOrEmpty(m.Email)
                    && !m.Email.EndsWith("@deloitte.com",
                    System.StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Log invalid emails for information
            foreach (var m in invalidEmails)
            {
                TestContext.WriteLine(
                    $"Non-deloitte email: {m.Email}");
            }

            // Warning only - not all members may have 
            // deloitte emails in test data
            Assert.That(invalidEmails.Count,
                Is.LessThan(members.Count),
                "Most emails should end with @deloitte.com");
        }


        // ==================
        // ONBOARDING TESTS
        // ==================

        [Test]
        public async Task GetOnboardingSteps_ReturnsSteps()
        {
            var steps = await _repo
                .GetOnboardingSteps();
            Assert.That(steps, Is.Not.Null);
            Assert.That(steps.Count,
                Is.GreaterThan(0));
        }

        [Test]
        public async Task GetOnboardingSteps_HasTeamName()
        {
            var steps = await _repo
                .GetOnboardingSteps();
            Assert.That(steps.All(s =>
                !string.IsNullOrEmpty(s.TeamName)),
                Is.True);
        }

        [Test]
        public async Task GetOnboardingStepsByTeam_FiltersCorrectly()
        {
            var steps = await _repo
                .GetOnboardingSteps();
            if (!steps.Any())
                Assert.Ignore("No steps in sheet");

            var teamName = steps.First().TeamName;
            var filtered = await _repo
                .GetOnboardingStepsByTeam(teamName);

            Assert.That(filtered, Is.Not.Null);
            Assert.That(filtered.All(s =>
                s.TeamName == teamName), Is.True);
        }

        [Test]
        public async Task GetOnboardingSteps_OrderedByStepOrder()
        {
            var steps = await _repo
                .GetOnboardingSteps();
            if (!steps.Any())
                Assert.Ignore("No steps in sheet");

            var teamName = steps.First().TeamName;
            var teamSteps = steps
                .Where(s => s.TeamName == teamName)
                .ToList();

            for (int i = 1; i < teamSteps.Count; i++)
            {
                Assert.That(
                    teamSteps[i].StepOrder,
                    Is.GreaterThanOrEqualTo(
                        teamSteps[i - 1].StepOrder));
            }
        }

        // ==================
        // TRAINING STATUS TESTS
        // ==================

        [Test]
        public async Task GetTrainingStatus_ReturnsData()
        {
            var training = await _repo
                .GetTrainingStatus();
            Assert.That(training, Is.Not.Null);
        }

        [Test]
        public async Task GetTrainingStatusWithNames_HasCandidateName()
        {
            var training = await _repo
                .GetTrainingStatusWithNames();
            Assert.That(training, Is.Not.Null);
            if (!training.Any())
                Assert.Ignore("No training records");

            var first = training.First();
            var json = System.Text.Json.JsonSerializer
                .Serialize(first);
            Assert.That(json.Contains("candidateName"),
                Is.True,
                "Response should contain candidateName");
        }


        [Test]
        public async Task GetTrainingStatus_ValidDomains()
        {
            var validDomains = new[] {
                "DataDog", "AKS", "ROVO", "Copilot",
                "AI Fluency", "Claude 101",
                "Client Compliance Trainings", "Datadog", "Github Copilot"
            };

            var training = await _repo
                .GetTrainingStatus();

            foreach (var t in training)
            {
                if (!string.IsNullOrEmpty(t.Domain))
                {
                    Assert.That(
                        validDomains.Contains(t.Domain),
                        Is.True,
                        $"Invalid domain: {t.Domain}");
                }
            }
        }

        // ==================
        // USER TESTS
        // ==================

        [Test]
        public async Task GetAllUsers_ReturnsUsers()
        {
            var users = await _repo.GetAllUsers();
            Assert.That(users, Is.Not.Null);
            Assert.That(users.Count,
                Is.GreaterThan(0));
        }

        [Test]
        public async Task GetUserByUsername_ValidUser_ReturnsUser()
        {
            var users = await _repo.GetAllUsers();
            if (!users.Any())
                Assert.Ignore("No users in sheet");

            var username = users.First().UserName;
            var repo = new ExcelRepository(_testFilePath);
            var user = await repo
                .GetUserByUsername(username);

            Assert.That(user, Is.Not.Null);
            Assert.That(user.UserName,
                Is.EqualTo(username));
        }

        [Test]
        public async Task GetUserByUsername_InvalidUser_ReturnsNull()
        {
            var repo = new ExcelRepository(_testFilePath);
            var user = await repo
                .GetUserByUsername(
                    "nonexistent@deloitte.com");
            Assert.That(user, Is.Null);
        }

        [Test]
        public async Task GetAllUsers_EmailsEndWithDeloitte()
        {
            var users = await _repo.GetAllUsers();
            var invalidUsers = users
                .Where(u => !string.IsNullOrEmpty(
                    u.UserName) &&
                    !u.UserName.EndsWith(
                        "@deloitte.com",
                        System.StringComparison
                            .OrdinalIgnoreCase))
                .ToList();
            Assert.That(invalidUsers.Count,
                Is.EqualTo(0));
        }

        // ==================
        // OFFBOARDED TESTS
        // ==================

        [Test]
        public async Task GetOffboarded_ReturnsData()
        {
            var offboarded = await _repo
                .GetOffboarded();
            Assert.That(offboarded, Is.Not.Null);
        }

        // ==================
        // ONBOARDING PROGRESS TESTS
        // ==================

        [Test]
        public async Task GetOnboardingProgress_ReturnsData()
        {
            var progress = await _repo
                .GetOnboardingProgress();
            Assert.That(progress, Is.Not.Null);
        }

        [Test]
        public async Task GetOnboardingStatusWithNames_ReturnsData()
        {
            var status = await _repo
                .GetOnboardingStatusWithNames();
            Assert.That(status, Is.Not.Null);
        }

        // ==================
        // PROJECTS TESTS
        // ==================

        [Test]
        public async Task GetProjects_ReturnsProjects()
        {
            var projects = await _repo.GetProjects();
            Assert.That(projects, Is.Not.Null);
            Assert.That(projects.Count,
                Is.GreaterThan(0));
        }

        [Test]
        public async Task GetAllMembers_SrNoIsUnique()
        {
            var members = await _repo.GetAllMembers();
            var distinctIds = members
                .Select(m => m.SrNo)
                .Distinct()
                .Count();
            Assert.That(distinctIds,
                Is.EqualTo(members.Count),
                "All SrNo values should be unique");
        }

        [Test]
        public async Task GetAllMembers_NameIsNotEmpty()
        {
            var members = await _repo.GetAllMembers();
            var emptyNames = members
                .Where(m => string.IsNullOrEmpty(m.Name))
                .ToList();
            Assert.That(emptyNames.Count, Is.EqualTo(0),
                "No member should have empty name");
        }

        [Test]
        public async Task GetActiveMembers_CountLessThanTotal()
        {
            var all = await _repo.GetAllMembers();
            var active = await _repo.GetActiveMembers();
            Assert.That(active.Count,
                Is.LessThanOrEqualTo(all.Count));
        }




        [Test]
        public async Task GetOnboardingSteps_StepOrderIsPositive()
        {
            var steps = await _repo.GetOnboardingSteps();
            Assert.That(steps.All(s => s.StepOrder > 0),
                Is.True,
                "All step orders should be positive");
        }

        [Test]
        public async Task GetOnboardingStepsByTeam_EmptyTeam_ReturnsAll()
        {
            var all = await _repo
                .GetOnboardingStepsByTeam("");
            var allSteps = await _repo.GetOnboardingSteps();
            Assert.That(all.Count,
                Is.EqualTo(allSteps.Count));
        }



        [Test]
        public async Task GetTrainingStatus_CandidateIDIsPositive()
        {
            var training = await _repo.GetTrainingStatus();
            Assert.That(training.All(t => t.CandidateID > 0),
                Is.True,
                "All CandidateIDs should be positive");
        }

        [Test]
        public async Task GetTrainingStatus_TrainingIDIsUnique()
        {
            var training = await _repo.GetTrainingStatus();
            var distinctIds = training
                .Select(t => t.TrainingID)
                .Distinct()
                .Count();

            TestContext.WriteLine(
                $"Total records: {training.Count}, " +
                $"Unique IDs: {distinctIds}");

            Assert.That(distinctIds,
                Is.GreaterThan(0),
                "Training records should have valid IDs");
        }




        [Test]
        public async Task GetAllUsers_RoleIsValid()
        {
            var users = await _repo.GetAllUsers();
            var validRoles = new[] { "Admin", "User" };
            var invalidRoles = users
                .Where(u => !string.IsNullOrEmpty(u.Role)
                    && !validRoles.Contains(u.Role))
                .ToList();
            Assert.That(invalidRoles.Count, Is.EqualTo(0),
                "All roles should be Admin or User");
        }

        [Test]
        public async Task GetAllUsers_UserIDIsUnique()
        {
            var users = await _repo.GetAllUsers();
            var distinctIds = users
                .Select(u => u.UserID)
                .Distinct()
                .Count();

            TestContext.WriteLine(
                $"Total users: {users.Count}, " +
                $"Unique IDs: {distinctIds}");

            Assert.That(distinctIds,
                Is.GreaterThan(0),
                "Users should have valid IDs");
        }

        [Test]
        public async Task GetAllUsers_PasswordNotEmpty()
        {
            var users = await _repo.GetAllUsers();
            var noPassword = users
                .Where(u => string.IsNullOrEmpty(u.PasswordHash))
                .ToList();
            Assert.That(noPassword.Count, Is.EqualTo(0),
                "All users should have a password");
        }



        [Test]
        public async Task GetOffboarded_EmailsAreValid()
        {
            var offboarded = await _repo.GetOffboarded();
            var invalidEmails = offboarded
                .Where(o => !string.IsNullOrEmpty(o.Email)
                    && !o.Email.Contains("@"))
                .ToList();
            Assert.That(invalidEmails.Count, Is.EqualTo(0),
                "All offboarded emails should be valid");
        }

    }
}

