using OfficeOpenXml;
using TeamTracker.API.Models;

namespace TeamTracker.API.Repositories
{
    public class ExcelRepository
    {
        private readonly string _filePath;
        private static readonly SemaphoreSlim _lock
            = new SemaphoreSlim(1, 1);

        public ExcelRepository(string filePath)
        {
            _filePath = filePath;
        }

        public async Task<User?> GetUserByUsername(
            string username)
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

                if (sheet == null) return null;

                for (int row = 2; row <= sheet
                    .Dimension.End.Row; row++)
                {
                    var uname = sheet.Cells[row, 2]
                        .Value?.ToString();
                    if (uname?.Trim() ==
                        username.Trim())
                    {
                        return new User
                        {
                            UserID = Convert.ToInt32(
                                sheet.Cells[row, 1]
                                .Value ?? 0),
                            UserName = uname,
                            PasswordHash = sheet
                                .Cells[row, 3]
                                .Value?.ToString(),
                            Role = sheet.Cells[row, 4]
                                .Value?.ToString(),
                            IsActive = sheet
                                .Cells[row, 5]
                                .Value?.ToString()
                                ?.ToLower() == "true"
                        };
                    }
                }
                return null;
            }
            finally
            {
                _lock.Release();
            }
        }
    }
}
