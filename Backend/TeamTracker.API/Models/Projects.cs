namespace TeamTracker.API.Models
{
    public class Projects
    {
        public string? Sno {  get; set; }
        public string? ProjectName { get; set; }
        public string? Status { get; set; }
        public string? TeamCount { get; set; }
        public string? DeloitteTeamLead { get; set; }
        public string? HasUSATeam { get; set; }
        public string? HasADMXTeam { get; set; }
        public string? ScrumMasterOffshore { get; set; }
        public string? ScrumMasterOnshore { get; set; }
        public string? ScrumMasterRelease { get; set; }
        public string? MeetingCadence { get; set; }
    }
}
