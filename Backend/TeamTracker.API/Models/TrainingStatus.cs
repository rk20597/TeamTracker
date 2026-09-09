namespace TeamTracker.API.Models
{
    public class TrainingStatus
    {
        public int TrainingID { get; set; }
        public int CandidateID { get; set; }
        public string? Domain { get; set; }
        public string? Status { get; set; }
        public string? DueDate { get; set; }
        public string? CompletedDate { get; set; }
    }
}
