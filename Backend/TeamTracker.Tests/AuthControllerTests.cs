using NUnit.Framework;
using HROnboarding.API.Models;

namespace TeamTracker.Tests
{
    [TestFixture]
    public class ValidationTests
    {
        // ==================
        // EMAIL VALIDATION
        // ==================

        [Test]
        [TestCase("test@deloitte.com", true)]
        [TestCase("admin@deloitte.com", true)]
        [TestCase("test@gmail.com", false)]
        [TestCase("test@yahoo.com", false)]
        [TestCase("", false)]
        [TestCase("", false)]
        public void EmailValidation_DeloitteDomain(
            string email, bool expected)
        {
            bool result = !string.IsNullOrEmpty(email)
                && email.EndsWith("@deloitte.com",
                System.StringComparison.OrdinalIgnoreCase);
            Assert.That(result, Is.EqualTo(expected));
        }

        // ==================
        // PASSWORD VALIDATION
        // ==================

        [Test]
        [TestCase("Admin@123", true)]
        [TestCase("User@123!", true)]
        [TestCase("password", false)]
        [TestCase("pass", false)]
        [TestCase("Pass1!", false)]
        [TestCase("", false)]
        public void PasswordValidation_Rules(
            string password, bool expected)
        {
            bool result = ValidatePassword(password);
            Assert.That(result, Is.EqualTo(expected));
        }

        private bool ValidatePassword(string password)
        {
            if (string.IsNullOrEmpty(password))
                return false;
            if (password.Length < 8) return false;
            if (!System.Text.RegularExpressions
                .Regex.IsMatch(password, "[A-Z]"))
                return false;
            if (!System.Text.RegularExpressions
                .Regex.IsMatch(password, "[0-9]"))
                return false;
            if (!System.Text.RegularExpressions
                .Regex.IsMatch(password, "[!@#$%^&*]"))
                return false;
            return true;
        }

        // ==================
        // MOBILE VALIDATION
        // ==================

        [Test]
        [TestCase("9999999999", true)]
        [TestCase("1234567890", true)]
        [TestCase("123456789", false)]
        [TestCase("12345678901", false)]
        [TestCase("", false)]
        [TestCase("", false)]
        public void MobileValidation_TenDigits(
            string mobile, bool expected)
        {
            bool result = !string.IsNullOrEmpty(mobile)
                && mobile.Length == 10
                && mobile.All(char.IsDigit);
            Assert.That(result, Is.EqualTo(expected));
        }

        // ==================
        // TRAINING STATUS VALIDATION
        // ==================

        [Test]
        [TestCase("NotStarted", true)]
        [TestCase("InProgress", true)]
        [TestCase("Completed", true)]
        [TestCase("completed", false)]
        [TestCase("", false)]
        [TestCase("", false)]
        public void TrainingStatus_ValidValues(
            string status, bool expected)
        {
            var validStatuses = new[] {
                "NotStarted", "InProgress", "Completed"
            };
            bool result = !string.IsNullOrEmpty(status)
                && validStatuses.Contains(status);
            Assert.That(result, Is.EqualTo(expected));
        }

        // ==================
        // ONBOARDING STEP VALIDATION
        // ==================

        [Test]
        [TestCase(1, true)]
        [TestCase(5, true)]
        [TestCase(0, false)]
        [TestCase(-1, false)]
        public void StepOrder_MustBePositive(
            int stepOrder, bool expected)
        {
            bool result = stepOrder > 0;
            Assert.That(result, Is.EqualTo(expected));
        }

        // ==================
        // ROLE VALIDATION
        // ==================

        [Test]
        [TestCase("Admin", true)]
        [TestCase("User", true)]
        [TestCase("admin", false)]
        [TestCase("Manager", false)]
        [TestCase("", false)]
        public void RoleValidation_ValidRoles(
            string role, bool expected)
        {
            var validRoles = new[] { "Admin", "User" };
            bool result = !string.IsNullOrEmpty(role)
                && validRoles.Contains(role);
            Assert.That(result, Is.EqualTo(expected));
        }
    }
}
