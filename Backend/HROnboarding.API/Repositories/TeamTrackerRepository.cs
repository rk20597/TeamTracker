using OfficeOpenXml;
using HROnboarding.API.Models;
using System.Security.Cryptography.X509Certificates;

namespace HROnboarding.API.Repositories
{
    public class TeamTrackerRepository
    {
        private readonly string _filePath;
        private static readonly SemaphoreSlim _lock
            = new SemaphoreSlim(1, 1);

        public TeamTrackerRepository(string filePath)
        {
            _filePath = filePath;
        }

        // ==================
        // TEAM MEMBERS
        // ==================

        public async Task<List<TeamMember>>
            GetAllMembers()
        {
            await _lock.WaitAsync();
            try
            {
                var members = new List<TeamMember>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "TeamMember")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return members;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var name = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(name))
                        continue;

                    members.Add(new TeamMember
                    {
                        SrNo = row - 1,
                        Name = name,
                        Email = sheet.Cells[row, 3]
                            .Value?.ToString(),
                        Location = sheet.Cells[row, 4]
                            .Value?.ToString(),
                        Status = sheet.Cells[row, 5]
                            .Value?.ToString(),
                        Expansion = sheet.Cells[row, 6]
                            .Value?.ToString(),
                        City = sheet.Cells[row, 7]
                            .Value?.ToString(),
                        LevelOriginal = sheet
                            .Cells[row, 8]
                            .Value?.ToString(),
                        Level = sheet.Cells[row, 9]
                            .Value?.ToString(),
                        ClientLevel = sheet
                            .Cells[row, 10]
                            .Value?.ToString(),
                        OnboardingDate = sheet
                            .Cells[row, 11]
                            .Value?.ToString(),
                        Joined = sheet.Cells[row, 12]
                            .Value?.ToString(),
                        JobFamily = sheet
                            .Cells[row, 13]
                            .Value?.ToString(),
                        GDLeader1 = sheet
                            .Cells[row, 14]
                            .Value?.ToString(),
                        GDLeader2 = sheet
                            .Cells[row, 15]
                            .Value?.ToString(),
                        GDLeader3 = sheet
                            .Cells[row, 16]
                            .Value?.ToString(),
                        Project = sheet.Cells[row, 17]
                            .Value?.ToString(),
                        ClientKTDelivery = sheet
                            .Cells[row, 18]
                            .Value?.ToString(),
                        ProjectDeliveryStarted = sheet
                            .Cells[row, 19]
                            .Value?.ToString(),
                        ToolAccessCompleted = sheet
                            .Cells[row, 20]
                            .Value?.ToString(),
                        Datadog = sheet.Cells[row, 21]
                            .Value?.ToString(),
                        AKSTrained = sheet
                            .Cells[row, 22]
                            .Value?.ToString(),
                        ROVO = sheet.Cells[row, 23]
                            .Value?.ToString(),
                        AIFluency = sheet
                            .Cells[row, 25]
                            .Value?.ToString(),
                        Claude101 = sheet
                            .Cells[row, 26]
                            .Value?.ToString(),
                        OtherAICertification = sheet
                            .Cells[row, 27]
                            .Value?.ToString(),
                        CoPilotTrained = sheet
                            .Cells[row, 28]
                            .Value?.ToString(),
                        ClientComplianceTrainings = sheet
                            .Cells[row, 29]
                            .Value?.ToString(),
                        SciFormaAccess = sheet
                            .Cells[row, 30]
                            .Value?.ToString(),
                        OffboardingDate = sheet
                            .Cells[row, 32]
                            .Value?.ToString(),
                        Comments = sheet
                            .Cells[row, 31]
                            .Value?.ToString(),
                        MobileNumber = sheet
                            .Cells[row, 33]
                            .Value?.ToString(),
                        Replacement = sheet
                            .Cells[row, 34]
                            .Value?.ToString()
                    });
                }
                return members;
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task<List<TeamMember>>
            GetActiveMembers()
        {
            var all = await GetAllMembers();
            return all.Where(m =>
                m.Status?.ToLower() == "active")
                .ToList();
        }

        public async Task<List<TeamMember>>
            GetInactiveMembers()
        {
            var all = await GetAllMembers();
            return all.Where(m =>
                m.Status?.ToLower() == "inactive")
                .ToList();
        }

        public async Task DeleteTeamMember(int srNo)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "TeamMember")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var id = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    if (id == srNo)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task UpdateTeamMember(TeamMember member)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook.Worksheets)
                {
                    Console.WriteLine("Sheet found: " + ws.Name);
                    if (ws.Name.Trim() == "TeamMember")
                    {
                        sheet = ws;
                        break;
                    }
                }

                Console.WriteLine("Sheet null: " + (sheet == null));

                if (sheet == null) return;
                if (sheet.Dimension == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var idVal = sheet.Cells[row, 1]
                        .Value?.ToString();
                    int id = 0;
                    int.TryParse(idVal, out id);
                    Console.WriteLine("Row:" + row +
                        " ID:" + id +
                        " Looking for:" + member.SrNo);
                    if (id == member.SrNo)
                    {
                        Console.WriteLine("Found row " + row);
                        sheet.Cells[row, 2].Value = member.Name;
                        sheet.Cells[row, 3].Value = member.Email;
                        sheet.Cells[row, 4].Value = member.Location;
                        sheet.Cells[row, 5].Value = member.Status;
                        sheet.Cells[row, 6].Value = member.Expansion;
                        sheet.Cells[row, 7].Value = member.City;
                        sheet.Cells[row, 8].Value = member.LevelOriginal;
                        sheet.Cells[row, 9].Value = member.Level;
                        sheet.Cells[row, 10].Value = member.ClientLevel;
                        sheet.Cells[row, 11].Value = member.OnboardingDate;
                        sheet.Cells[row, 12].Value = member.Joined;
                        sheet.Cells[row, 13].Value = member.JobFamily;
                        sheet.Cells[row, 14].Value = member.GDLeader1;
                        sheet.Cells[row, 15].Value = member.GDLeader2;
                        sheet.Cells[row, 16].Value = member.GDLeader3;
                        sheet.Cells[row, 17].Value = member.Project;
                        sheet.Cells[row, 18].Value = member.ClientKTDelivery;
                        sheet.Cells[row, 19].Value = member.ProjectDeliveryStarted;
                        sheet.Cells[row, 20].Value = member.ToolAccessCompleted;
                        sheet.Cells[row, 21].Value = member.Datadog;
                        sheet.Cells[row, 22].Value = member.AKSTrained;
                        sheet.Cells[row, 23].Value = member.ROVO;
                        sheet.Cells[row, 24].Value = member.ClaudeCodeDevDay;
                        sheet.Cells[row, 25].Value = member.AIFluency;
                        sheet.Cells[row, 26].Value = member.Claude101;
                        sheet.Cells[row, 27].Value = member.OtherAICertification;
                        sheet.Cells[row, 28].Value = member.CoPilotTrained;
                        sheet.Cells[row, 29].Value = member.ClientComplianceTrainings;
                        sheet.Cells[row, 30].Value = member.SciFormaAccess;
                        sheet.Cells[row, 31].Value = member.Comments;
                        sheet.Cells[row, 32].Value = member.OffboardingDate;
                        sheet.Cells[row, 33].Value = member.MobileNumber;
                        sheet.Cells[row, 34].Value = member.Replacement;
                        
                        break;
                    }
                }
                Console.WriteLine("Saving...");   
                await package.SaveAsync();
                Console.WriteLine("Saved!");
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task AddTeamMember(TeamMember member)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "TeamMember")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                int newRow = sheet.Dimension.End.Row + 1;
                sheet.Cells[newRow, 1].Value = newRow - 1;
                sheet.Cells[newRow, 2].Value = member.Name;
                sheet.Cells[newRow, 3].Value = member.Email;
                sheet.Cells[newRow, 4].Value = member.Location;
                sheet.Cells[newRow, 5].Value = member.Status;
                sheet.Cells[newRow, 6].Value = member.Expansion;
                sheet.Cells[newRow, 7].Value = member.City;
                sheet.Cells[newRow, 8].Value = member.LevelOriginal;
                sheet.Cells[newRow, 9].Value = member.Level;
                sheet.Cells[newRow, 10].Value = member.ClientLevel;
                sheet.Cells[newRow, 31].Value = member.OnboardingDate;
                sheet.Cells[newRow, 12].Value = member.Joined;
                sheet.Cells[newRow, 13].Value = member.JobFamily;
                sheet.Cells[newRow, 14].Value = member.GDLeader1;
                sheet.Cells[newRow, 15].Value = member.GDLeader2;
                sheet.Cells[newRow, 16].Value = member.GDLeader3;
                sheet.Cells[newRow, 17].Value = member.Project;
                sheet.Cells[newRow, 18].Value = member.ClientKTDelivery;
                sheet.Cells[newRow, 19].Value = member.ProjectDeliveryStarted;
                sheet.Cells[newRow, 20].Value = member.ToolAccessCompleted;
                sheet.Cells[newRow, 21].Value = member.Datadog;
                sheet.Cells[newRow, 22].Value = member.AKSTrained;
                sheet.Cells[newRow, 23].Value = member.ROVO;
                sheet.Cells[newRow, 24].Value = member.ClaudeCodeDevDay;
                sheet.Cells[newRow, 25].Value = member.AIFluency;
                sheet.Cells[newRow, 26].Value = member.Claude101;
                sheet.Cells[newRow, 27].Value = member.OtherAICertification;
                sheet.Cells[newRow, 28].Value = member.CoPilotTrained;
                sheet.Cells[newRow, 29].Value = member.ClientComplianceTrainings;
                sheet.Cells[newRow, 30].Value = member.SciFormaAccess;
                sheet.Cells[newRow, 31].Value = member.Comments;
                sheet.Cells[newRow, 32].Value = member.OffboardingDate;
                sheet.Cells[newRow, 33].Value = member.MobileNumber;
                sheet.Cells[newRow, 34].Value = member.Replacement;
                

                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }



        // ==================
        // ONBOARDING STEPS
        // ==================

        public async Task<List<OnboardingStep>>
            GetOnboardingSteps()
        {
            await _lock.WaitAsync();
            try
            {
                var steps = new List<OnboardingStep>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() ==
                        "OnboardingSteps")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return steps;

                for (int row = 2; row <= sheet
    .Dimension.End.Row; row++)
                {
                    var stepName = sheet.Cells[row, 3]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(stepName))
                        continue;

                    var stepIdVal = sheet.Cells[row, 1]
                        .Value?.ToString();
                    int stepId = 0;
                    int.TryParse(stepIdVal, out stepId);

                    var stepOrderVal = sheet.Cells[row, 4]
                        .Value?.ToString();
                    int stepOrder = 0;
                    int.TryParse(stepOrderVal, out stepOrder);

                    steps.Add(new OnboardingStep
                    {
                        StepID = stepId,
                        TeamName = sheet.Cells[row, 2]
                            .Value?.ToString(),
                        StepName = stepName,
                        StepOrder = stepOrder,
                        Description = sheet.Cells[row, 5]
                            .Value?.ToString()
                    });
                }
                return steps;

            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task AddOnboardingStep(
            OnboardingStep step)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() ==
                        "OnboardingSteps")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                int newRow = sheet.Dimension
                    .End.Row + 1;
                sheet.Cells[newRow, 1].Value =
                    newRow - 1;
                sheet.Cells[newRow, 2].Value =
                    step.TeamName;
                sheet.Cells[newRow, 3].Value =
                    step.StepName;
                sheet.Cells[newRow, 4].Value =
                    step.StepOrder;
                sheet.Cells[newRow, 5].Value =
                    step.Description;

                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task UpdateOnboardingStep(
            OnboardingStep step)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() ==
                        "OnboardingSteps")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var id = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    if (id == step.StepID)
                    {
                        sheet.Cells[row, 2].Value =
                            step.TeamName;
                        sheet.Cells[row, 3].Value =
                            step.StepName;
                        sheet.Cells[row, 4].Value =
                            step.StepOrder;
                        sheet.Cells[row, 5].Value =
                            step.Description;
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task DeleteOnboardingStep(
            int stepId)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() ==
                        "OnboardingSteps")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var id = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    if (id == stepId)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task<List<OnboardingProgress>>
    GetOnboardingProgress()
        {
            await _lock.WaitAsync();
            try
            {
                var progress =
                    new List<OnboardingProgress>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() ==
                        "OnboardingProgress")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return progress;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    progress.Add(new OnboardingProgress
                    {
                        ProgressID = Convert.ToInt32(
                            sheet.Cells[row, 1]
                            .Value ?? 0),
                        CandidateID = Convert.ToInt32(
                            sheet.Cells[row, 2]
                            .Value ?? 0),
                        StepID = Convert.ToInt32(
                            sheet.Cells[row, 3]
                            .Value ?? 0),
                        CompletedDate = sheet
                            .Cells[row, 5]
                            .Value?.ToString(),
                        Status = sheet.Cells[row, 4]
                            .Value?.ToString()
                    });
                }
                return progress;
            }
            finally
            {
                _lock.Release();
            }
        }

        // DELETE ONBOARDING PROGRESS
        public async Task DeleteProgress(
            int candidateId, int stepId)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "OnboardingProgress")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var candId = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    var stId = Convert.ToInt32(
                        sheet.Cells[row, 2].Value ?? 0);
                    if (candId == candidateId &&
                        stId == stepId)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }



        // ==================
        // TRAINING LINKS
        // ==================

        public async Task<List<TrainingCourse>>
            GetTrainingCourses(string sheetName)
        {
            await _lock.WaitAsync();
            try
            {
                var courses = new List<TrainingCourse>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == sheetName)
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return courses;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var title = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(title))
                        continue;
                    courses.Add(new TrainingCourse
                    {
                        Platform = sheet.Cells[row, 1]
                            .Value?.ToString(),
                        Title = title,
                        Link = sheet.Cells[row, 3]
                            .Value?.ToString(),
                        EstTime = sheet.Cells[row, 4]
                            .Value?.ToString(),
                        Remark = sheet.Cells[row, 5]
                            .Value?.ToString(),
                        Domain = sheetName
                    });
                }
                return courses;
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task<List<TrainingLink>>
            GetTrainingLinks(string sheetName)
        {
            await _lock.WaitAsync();
            try
            {
                var links = new List<TrainingLink>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == sheetName)
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return links;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var topic = sheet.Cells[row, 1]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(topic))
                        continue;
                    links.Add(new TrainingLink
                    {
                        Topic = topic,
                        Link = sheet.Cells[row, 2]
                            .Value?.ToString(),
                        Domain = sheetName
                    });
                }
                return links;
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task<List<MandatoryTrainings>>
            GetMandatoryTrainings()
        {
            await _lock.WaitAsync();
            try
            {
                var trainings =
                    new List<MandatoryTrainings>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Madatory Trainings")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return trainings;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var jobFamily = sheet.Cells[row, 1]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(jobFamily))
                        continue;
                    trainings.Add(new MandatoryTrainings
                    {
                        JobFamily = jobFamily,
                        MandatoryTraining = sheet
                            .Cells[row, 2]
                            .Value?.ToString(),
                        POC = sheet.Cells[row, 3]
                            .Value?.ToString()
                    });
                }
                return trainings;
            }
            finally
            {
                _lock.Release();
            }
        }

        // DELETE TRAINING LINK (Topic/Link format)
        public async Task DeleteTrainingLink(
            string sheetName, string topic)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == sheetName)
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var t = sheet.Cells[row, 1]
                        .Value?.ToString();
                    if (t == topic)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        // DELETE TRAINING COURSE (Platform/Title format)
        public async Task DeleteTrainingCourse(
            string sheetName, string title)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == sheetName)
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var t = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (t == title)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }



        // ==================
        // PROJECTS
        // ==================

        public async Task<List<Projects>> GetProjects()
        {
            await _lock.WaitAsync();
            try
            {
                var projects = new List<Projects>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim()
                        .Contains("Projects"))
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return projects;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var name = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(name))
                        continue;
                    projects.Add(new Projects
                    {
                        Sno = sheet.Cells[row, 1]
                            .Value?.ToString(),
                        ProjectName = name,
                        Status = sheet.Cells[row, 3]
                            .Value?.ToString(),
                        TeamCount = sheet.Cells[row, 4]
                            .Value?.ToString(),
                        DeloitteTeamLead = sheet
                            .Cells[row, 5]
                            .Value?.ToString(),
                        HasUSATeam = sheet.Cells[row, 6]
                            .Value?.ToString(),
                        HasADMXTeam = sheet
                            .Cells[row, 7]
                            .Value?.ToString(),
                        ScrumMasterOffshore = sheet
                            .Cells[row, 8]
                            .Value?.ToString(),
                        ScrumMasterOnshore = sheet
                            .Cells[row, 9]
                            .Value?.ToString(),
                        ScrumMasterRelease = sheet
                            .Cells[row, 10]
                            .Value?.ToString(),
                        MeetingCadence = sheet
                            .Cells[row, 11]
                            .Value?.ToString()
                    });
                }
                return projects;
            }
            finally
            {
                _lock.Release();
            }
        }

        // DELETE PROJECT
        public async Task DeleteProject(string projectName)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim().Contains("Project"))
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var name = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (name == projectName)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }


        // ==================
        // GD LEAD
        // ==================

        public async Task<List<GDLead>> GetGDLeadData()
        {
            await _lock.WaitAsync();
            try
            {
                var leads = new List<GDLead>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "GDLead")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return leads;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var skillset = sheet.Cells[row, 1]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(skillset))
                        continue;
                    leads.Add(new GDLead
                    {
                        SkillSet = skillset,
                        GDLeadName = sheet
                            .Cells[row, 2]
                            .Value?.ToString(),
                        TotalResources = sheet
                            .Cells[row, 3]
                            .Value?.ToString(),
                        AdditionalTools = sheet
                            .Cells[row, 4]
                            .Value?.ToString(),
                        KeyHighlights = sheet
                            .Cells[row, 5]
                            .Value?.ToString()
                    });
                }
                return leads;
            }
            finally
            {
                _lock.Release();
            }
        }

        // DELETE GD LEAD
        public async Task DeleteGDLead(string skillSet)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "GDLead")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var skill = sheet.Cells[row, 1]
                        .Value?.ToString();
                    if (skill == skillSet)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }


        // ==================
        // RAW SHEET DATA
        // For pivot tables
        // ==================

        public async Task<List<Dictionary
            <string, string>>>
            GetRawSheetData(string sheetName)
        {
            await _lock.WaitAsync();
            try
            {
                var data = new List<Dictionary
                    <string, string>>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == sheetName)
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return data;

                var headers = new List<string>();
                for (int col = 1; col <= sheet
                    .Dimension.End.Column; col++)
                {
                    headers.Add(sheet.Cells[1, col]
                        .Value?.ToString() ??
                        $"Col{col}");
                }

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var rowData = new Dictionary
                        <string, string>();
                    for (int col = 1; col <=
                        headers.Count; col++)
                    {
                        rowData[headers[col - 1]] =
                            sheet.Cells[row, col]
                            .Value?.ToString() ?? "";
                    }
                    data.Add(rowData);
                }
                return data;
            }
            finally
            {
                _lock.Release();
            }
        }

        // GET TRAINING STATUS
        public async Task<List<TrainingStatus>> GetTrainingStatus()
        {
            await _lock.WaitAsync();
            try
            {
                var list = new List<TrainingStatus>();
                using var package = new ExcelPackage(new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook.Worksheets)
                {
                    if (ws.Name.Trim() == "Training Status")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return list;

                for (int row = 2; row <= sheet.Dimension.End.Row; row++)
                {
                    var domain = sheet.Cells[row, 3].Value?.ToString();
                    if (string.IsNullOrEmpty(domain))
                        continue;

                    list.Add(new TrainingStatus
                    {
                        TrainingID = Convert.ToInt32(sheet.Cells[row, 1].Value ?? 0),
                        CandidateID = Convert.ToInt32(sheet.Cells[row, 2].Value ?? 0),
                        Domain = domain,
                        Status = sheet.Cells[row, 4].Value?.ToString(),
                        DueDate = sheet.Cells[row, 5].Value?.ToString(),
                        CompletedDate = sheet.Cells[row, 6].Value?.ToString(),
                    });
                }
                return list;
            }
            finally
            { 
                _lock.Release();
            }  
       
        }

        public async Task<List<object>> GetTrainingStatusWithNames()
        {
            var training = await GetTrainingStatus();
            var members = await GetAllMembers();

            var result = training.Select(t => {
                var member = members.FirstOrDefault(
                    m => m.SrNo == t.CandidateID);
                return (object)new
                {
                    trainingID = t.TrainingID,
                    candidateID = t.CandidateID,
                    candidateName = member?.Name ??
                        "Unknown",
                    domain = t.Domain,
                    status = t.Status,
                    dueDate = t.DueDate,
                    completedDate = t.CompletedDate,
                    isOverdue = !string.IsNullOrEmpty(
                        t.DueDate) &&
                        t.Status != "Completed" &&
                        DateTime.TryParse(t.DueDate,
                            out DateTime due) &&
                        due < DateTime.Now
                };
            }).ToList();

            return result;
        }


        // ADD TRAINING STATUS
        public async Task AddTrainingStatus(TrainingStatus training)
        { 
            await _lock.WaitAsync();
            try
            { 
                using var package = new ExcelPackage(new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook.Worksheets)
                {
                    if (ws.Name.Trim() == "Training Status")
                    { 
                        sheet = ws; 
                        break;
                    }
                }

                if (sheet == null) return;

                int newRow = (sheet.Dimension?.End?.Row ?? 1) + 1;
                sheet.Cells[newRow, 1].Value = newRow - 1;
                sheet.Cells[newRow, 2].Value = training.CandidateID;
                sheet.Cells[newRow, 3].Value = training.Domain;
                sheet.Cells[newRow, 4].Value = training.Status;
                sheet.Cells[newRow, 5].Value = training.DueDate;
                sheet.Cells[newRow, 6].Value = training.CompletedDate;

                await package.SaveAsync();
            }
            finally { _lock.Release(); }
        }

        //UPDATE TRAINING STATUS
        public async Task UpdateTrainingStatus(TrainingStatus training)
        {
            await _lock.WaitAsync();

            try
            {
                using var package = new ExcelPackage(new FileInfo(_filePath));
                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook.Worksheets)
                {
                    if (ws.Name.Trim() == "Training Status")
                    {
                        sheet = ws;
                        break;
                    }
                }
                
                if (sheet == null) return;
                for (int row = 2; row <= sheet
    .Dimension.End.Row; row++)
                {
                    var idVal = sheet.Cells[row, 1]
                        .Value?.ToString();
                    int id = 0;
                    int.TryParse(idVal, out id);
                    if (id == training.TrainingID)
                    {
                        sheet.Cells[row, 3].Value =
                            training.Domain;
                        sheet.Cells[row, 4].Value =
                            training.Status;
                        sheet.Cells[row, 5].Value =
                            training.DueDate;
                        sheet.Cells[row, 6].Value =
                            training.CompletedDate;
                        break;
                    }
                }
                Console.WriteLine("Saving...");
                await package.SaveAsync();
                Console.WriteLine("Saved!");

            }
            finally
            { 
                _lock.Release();
            }
        }

        // DELETE TRAINING STATUS
        public async Task DeleteTrainingStatus(int trainingId)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Training Status")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var id = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    if (id == trainingId)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }


        // GET OFFBOARDED
        public async Task<List<Offboarded>>
            GetOffboarded()
        {
            await _lock.WaitAsync();
            try
            {
                var list = new List<Offboarded>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Offboarding")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return list;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var name = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(name))
                        continue;
                    list.Add(new Offboarded
                    {
                        CandidateID = Convert.ToInt32(
                            sheet.Cells[row, 1].Value ?? 0),
                        Name = name,
                        Email = sheet.Cells[row, 3]
                            .Value?.ToString(),
                        Role = sheet.Cells[row, 4]
                            .Value?.ToString(),
                        Reason = sheet.Cells[row, 5]
                            .Value?.ToString()
                    });
                }
                return list;
            }
            finally
            {
                _lock.Release();
            }
        }

        // ADD OFFBOARDED
        public async Task AddOffboarded(Offboarded member)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Offboarding")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                int newRow = sheet.Dimension.End.Row + 1;
                sheet.Cells[newRow, 1].Value =
                    member.CandidateID;
                sheet.Cells[newRow, 2].Value = member.Name;
                sheet.Cells[newRow, 3].Value = member.Email;
                sheet.Cells[newRow, 4].Value = member.Role;
                sheet.Cells[newRow, 5].Value = member.Reason;

                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        // GET ONBOARDING STEPS BY TEAM
        public async Task<List<OnboardingStep>>
            GetOnboardingStepsByTeam(string teamName)
        {
            var all = await GetOnboardingSteps();
            if (string.IsNullOrEmpty(teamName))
                return all;
            return all.Where(s =>
                s.TeamName == teamName).ToList();
        }

        // GET PROGRESS BY CANDIDATE
        public async Task<List<OnboardingProgress>>
            GetProgressByCandidate(int candidateId)
        {
            var all = await GetOnboardingProgress();
            return all.Where(p =>
                p.CandidateID == candidateId).ToList();
        }

        // UPDATE PROGRESS
        public async Task UpdateProgress(
            OnboardingProgress progress)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "OnboardingProgress")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var candId = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    var stepId = Convert.ToInt32(
                        sheet.Cells[row, 2].Value ?? 0);
                    if (candId == progress.CandidateID &&
                        stepId == progress.StepID)
                    {
                        sheet.Cells[row, 3].Value =
                            progress.Status;
                        sheet.Cells[row, 4].Value =
                            progress.CompletedDate;
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }

        }

        public async Task AddOrUpdateProgress(
    OnboardingProgress progress)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    Console.WriteLine("Sheet: " + ws.Name);
                    if (ws.Name.Trim() == "OnboardingProgress")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null)
                {
                    Console.WriteLine("Sheet not found!");
                    return;
                }

                bool found = false;

                if (sheet.Dimension != null)
                {
                    for (int row = 2; row <= sheet
                        .Dimension.End.Row; row++)
                    {
                        var candIdVal = sheet.Cells[row, 2]
                            .Value?.ToString();
                        var stepIdVal = sheet.Cells[row, 3]
                            .Value?.ToString();
                        int candId = 0;
                        int stepId = 0;
                        int.TryParse(candIdVal, out candId);
                        int.TryParse(stepIdVal, out stepId);

                        if (candId == progress.CandidateID &&
                            stepId == progress.StepID)
                        {
                            sheet.Cells[row, 4].Value =
                                progress.Status;
                            sheet.Cells[row, 5].Value =
                                progress.CompletedDate;
                            found = true;
                            Console.WriteLine("Updated row " + row);
                            break;
                        }
                    }
                }

                if (!found)
                {
                    int newRow = sheet.Dimension == null ?
                        2 : sheet.Dimension.End.Row + 1;

                    int maxId = 0;
                    if (sheet.Dimension != null)
                    {
                        for (int r = 2; r <= sheet
                            .Dimension.End.Row; r++)
                        {
                            var idVal = sheet.Cells[r, 1]
                                .Value?.ToString();
                            int id = 0;
                            int.TryParse(idVal, out id);
                            if (id > maxId) maxId = id;
                        }
                    }

                    sheet.Cells[newRow, 1].Value = maxId + 1;
                    sheet.Cells[newRow, 2].Value =
                        progress.CandidateID;
                    sheet.Cells[newRow, 3].Value =
                        progress.StepID;
                    sheet.Cells[newRow, 4].Value =
                        progress.Status;
                    sheet.Cells[newRow, 5].Value =
                        progress.CompletedDate;
                    Console.WriteLine("Added new row " + newRow);
                }

                await package.SaveAsync();
                Console.WriteLine("Saved successfully!");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Save error: " +
                    saveEx.Message);
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task<List<object>>
    GetOnboardingStatusWithNames()
        {
            var members = await GetAllMembers();
            var steps = await GetOnboardingSteps();
            var progress = await GetOnboardingProgress();

            var result = new List<object>();

            foreach (var prog in progress)
            {
                var member = members.FirstOrDefault(
                    m => m.SrNo == prog.CandidateID);
                var step = steps.FirstOrDefault(
                    s => s.StepID == prog.StepID);

                if (member == null || step == null)
                    continue;

                result.Add((object)new
                {
                    progressID = prog.ProgressID,
                    candidateID = prog.CandidateID,
                    candidateName = member.Name,
                    email = member.Email,
                    jobFamily = member.JobFamily,
                    project = member.Project,
                    stepID = prog.StepID,
                    stepName = step.StepName,
                    stepOrder = step.StepOrder,
                    teamName = step.TeamName,
                    status = prog.Status ?? "NotStarted",
                    completedDate =
                        prog.CompletedDate ?? ""
                });
            }
            return result;
        }


        // ADD OR UPDATE PROGRESS
        public async Task AddOrUpdateProgresswithNames(
            OnboardingProgress progress)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "OnboardingProgress")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                bool found = false;

                if (sheet.Dimension != null)
                {
                    for (int row = 2; row <= sheet
                        .Dimension.End.Row; row++)
                    {
                        var candIdVal = sheet.Cells[row, 2]
                            .Value?.ToString();
                        var stepIdVal = sheet.Cells[row, 3]
                            .Value?.ToString();
                        int candId = 0;
                        int stepId = 0;
                        int.TryParse(candIdVal, out candId);
                        int.TryParse(stepIdVal, out stepId);

                        if (candId == progress.CandidateID &&
                            stepId == progress.StepID)
                        {
                            sheet.Cells[row, 4].Value =
                                progress.Status;
                            sheet.Cells[row, 5].Value =
                                progress.CompletedDate;
                            found = true;
                            break;
                        }
                    }
                }

                if (!found)
                {
                    int newRow = sheet.Dimension == null ?
                        2 : sheet.Dimension.End.Row + 1;

                    int maxId = 0;
                    if (sheet.Dimension != null)
                    {
                        for (int r = 2; r <= sheet
                            .Dimension.End.Row; r++)
                        {
                            var idVal = sheet.Cells[r, 1]
                                .Value?.ToString();
                            int id = 0;
                            int.TryParse(idVal, out id);
                            if (id > maxId) maxId = id;
                        }
                    }

                    sheet.Cells[newRow, 1].Value = maxId + 1;
                    sheet.Cells[newRow, 2].Value =
                        progress.CandidateID;
                    sheet.Cells[newRow, 3].Value =
                        progress.StepID;
                    sheet.Cells[newRow, 4].Value =
                        progress.Status;
                    sheet.Cells[newRow, 5].Value =
                        progress.CompletedDate;
                }

                await package.SaveAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task DeleteProgresswithNames(
    int candidateId, int stepId)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "OnboardingProgress")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;
                if (sheet.Dimension == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var candIdVal = sheet.Cells[row, 2]
                        .Value?.ToString();
                    var stepIdVal = sheet.Cells[row, 3]
                        .Value?.ToString();
                    int candId = 0;
                    int stepId2 = 0;
                    int.TryParse(candIdVal, out candId);
                    int.TryParse(stepIdVal, out stepId2);

                    if (candId == candidateId &&
                        stepId2 == stepId)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }





        // DELETE OFFBOARDED
        public async Task DeleteOffboarded(int candidateId)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Offboarding")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var id = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    if (id == candidateId)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        // GET ALL USERS
        public async Task<List<User>> GetAllUsers()
        {
            await _lock.WaitAsync();
            try
            {
                var users = new List<User>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Users")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return users;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var username = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(username))
                        continue;
                    users.Add(new User
                    {
                        UserID = Convert.ToInt32(
                            sheet.Cells[row, 1].Value ?? 0),
                        UserName = username,
                        PasswordHash = sheet.Cells[row, 3]
                            .Value?.ToString(),
                        Role = sheet.Cells[row, 4]
                            .Value?.ToString(),
                        IsActive = sheet.Cells[row, 5]
                            .Value?.ToString()
                            ?.ToLower() == "true"
                    });
                }
                return users;
            }
            finally
            {
                _lock.Release();
            }
        }

        // ADD USER
        public async Task AddUser(User user)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Users")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                int newRow = sheet.Dimension.End.Row + 1;
                sheet.Cells[newRow, 1].Value = newRow - 1;
                sheet.Cells[newRow, 2].Value = user.UserName;
                sheet.Cells[newRow, 3].Value =
                    user.PasswordHash;
                sheet.Cells[newRow, 4].Value = user.Role;
                sheet.Cells[newRow, 5].Value =
                    user.IsActive ? "TRUE" : "FALSE";

                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        // UPDATE USER
        public async Task UpdateUser(User user)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Users")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var id = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    if (id == user.UserID)
                    {
                        sheet.Cells[row, 2].Value =
                            user.UserName;
                        if (!string.IsNullOrEmpty(user.PasswordHash))
                        {
                            sheet.Cells[row, 3].Value =
                            user.PasswordHash;
                        }
                        sheet.Cells[row, 4].Value =
                            user.Role;
                        sheet.Cells[row, 5].Value =
                            user.IsActive ? "TRUE" : "FALSE";
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        // DELETE USER
        public async Task DeleteUser(int userId)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "Users")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var id = Convert.ToInt32(
                        sheet.Cells[row, 1].Value ?? 0);
                    if (id == userId)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task<List<PRSD>> GetPRSD()
        {
            await _lock.WaitAsync();
            try
            {
                var list = new List<PRSD>();
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "PRSD")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return list;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var lob = sheet.Cells[row, 1]
                        .Value?.ToString();
                    if (string.IsNullOrEmpty(lob))
                        continue;
                    list.Add(new PRSD
                    {
                        LOB = lob,
                        TrackName = sheet.Cells[row, 2]
                            .Value?.ToString(),
                        ProdSupportApplicable = sheet
                            .Cells[row, 3].Value?.ToString(),
                        ProdSupport = sheet.Cells[row, 4]
                            .Value?.ToString(),
                        ProdSupportConfidenceLevel = sheet
                            .Cells[row, 5].Value?.ToString(),
                        ProdSupportETA = sheet.Cells[row, 6]
                            .Value?.ToString(),
                        ProdSupportChallenges = sheet
                            .Cells[row, 7].Value?.ToString(),
                        TotalProdSupportCount = sheet
                            .Cells[row, 8].Value?.ToString(),
                        TentetiveETA = sheet.Cells[row, 9]
                            .Value?.ToString(),
                        ResourceName = sheet.Cells[row, 10]
                            .Value?.ToString(),
                        ContactNumber = sheet.Cells[row, 11]
                            .Value?.ToString(),
                        ReleaseSupportApplicable = sheet
                            .Cells[row, 12].Value?.ToString(),
                        ReleaseSupport = sheet.Cells[row, 13]
                            .Value?.ToString(),
                        TotalReleaseSupport = sheet
                            .Cells[row, 14].Value?.ToString()
                    });
                }
                return list;
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task AddPRSD(PRSD prsd)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "PRSD")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                int newRow = sheet.Dimension.End.Row + 1;
                sheet.Cells[newRow, 1].Value = prsd.LOB;
                sheet.Cells[newRow, 2].Value = prsd.TrackName;
                sheet.Cells[newRow, 3].Value =
                    prsd.ProdSupportApplicable;
                sheet.Cells[newRow, 4].Value =
                    prsd.ProdSupport;
                sheet.Cells[newRow, 5].Value =
                    prsd.ProdSupportConfidenceLevel;
                sheet.Cells[newRow, 6].Value =
                    prsd.ProdSupportETA;
                sheet.Cells[newRow, 7].Value =
                    prsd.ProdSupportChallenges;
                sheet.Cells[newRow, 8].Value =
                    prsd.TotalProdSupportCount;
                sheet.Cells[newRow, 9].Value =
                    prsd.TentetiveETA;
                sheet.Cells[newRow, 10].Value =
                    prsd.ResourceName;
                sheet.Cells[newRow, 11].Value =
                    prsd.ContactNumber;
                sheet.Cells[newRow, 12].Value =
                    prsd.ReleaseSupportApplicable;
                sheet.Cells[newRow, 13].Value =
                    prsd.ReleaseSupport;
                sheet.Cells[newRow, 14].Value =
                    prsd.TotalReleaseSupport;

                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task UpdatePRSD(PRSD prsd)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "PRSD")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var lob = sheet.Cells[row, 1]
                        .Value?.ToString();
                    var track = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (lob == prsd.LOB &&
                        track == prsd.TrackName)
                    {
                        sheet.Cells[row, 3].Value =
                            prsd.ProdSupportApplicable;
                        sheet.Cells[row, 4].Value =
                            prsd.ProdSupport;
                        sheet.Cells[row, 5].Value =
                            prsd.ProdSupportConfidenceLevel;
                        sheet.Cells[row, 6].Value =
                            prsd.ProdSupportETA;
                        sheet.Cells[row, 7].Value =
                            prsd.ProdSupportChallenges;
                        sheet.Cells[row, 8].Value =
                            prsd.TotalProdSupportCount;
                        sheet.Cells[row, 9].Value =
                            prsd.TentetiveETA;
                        sheet.Cells[row, 10].Value =
                            prsd.ResourceName;
                        sheet.Cells[row, 11].Value =
                            prsd.ContactNumber;
                        sheet.Cells[row, 12].Value =
                            prsd.ReleaseSupportApplicable;
                        sheet.Cells[row, 13].Value =
                            prsd.ReleaseSupport;
                        sheet.Cells[row, 14].Value =
                            prsd.TotalReleaseSupport;
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task DeletePRSD(
            string lob, string trackName)
        {
            await _lock.WaitAsync();
            try
            {
                using var package = new ExcelPackage(
                    new FileInfo(_filePath));

                ExcelWorksheet? sheet = null;
                foreach (var ws in package.Workbook
                    .Worksheets)
                {
                    if (ws.Name.Trim() == "PRSD")
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null) return;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var l = sheet.Cells[row, 1]
                        .Value?.ToString();
                    var t = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (l == lob && t == trackName)
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                await package.SaveAsync();
            }
            finally
            {
                _lock.Release();
            }
        }




    }
}
