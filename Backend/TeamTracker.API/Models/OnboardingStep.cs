namespace TeamTracker.API.Models
{
    public class OnboardingStep
    {
        public int StepID { get; set; }
        public string? TeamName { get; set; }
        public string? StepName { get; set; }
        public int StepOrder { get; set; }
        public string? Description { get; set; }

    }
}
