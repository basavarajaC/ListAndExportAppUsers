using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using OfficeOpenXml;
using AuthenticationException = System.Security.Authentication.AuthenticationException;

namespace ListAndExportAppUsers
{
    public class Program
    {
        private static GraphServiceClient? _graphClient;
        private static string? _accessToken;
        //STAGING
       // private static readonly string? ClientId = "<>";
       // private static readonly string? Tenant = "<>";
       // private static readonly string? ClientSecret = "<>";
       // private static readonly string? ObjectId = "<>";
       // private static readonly string? AzureServicePrinciple = "<>";
        
        //PRODUCTION
        private static readonly string? ClientId = "<>";
        private static readonly string? Tenant = "<>";
        private static readonly string? ClientSecret = "<>";
        private static readonly string? ObjectId = "<>";
        private static readonly string? AzureServicePrinciple = "<>";

        private static async Task Main()
        {
            await ExportUsersAssignedToApp();
        }

        private static async Task<string> GetAccessToken()
        {
            var scopes = new List<string> { "https://graph.microsoft.com/.default" };

            var msalClient = ConfidentialClientApplicationBuilder
                .Create(ClientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, Tenant)
                .WithClientSecret(ClientSecret)
                .Build();

            try
            {
                var token = await msalClient.AcquireTokenForClient(scopes).ExecuteAsync();
                return token.AccessToken;
            }
            catch (Exception ex)
            {
                throw new AuthenticationException("Issue on getting access token", ex);
            }
        }

        private static GraphServiceClient? InitiateGraphServiceClient(string accessToken)
        {
            var authenticationProvider = new DelegateAuthenticationProvider(
                requestMessage =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    return Task.FromResult(0);
                });
            return new GraphServiceClient(authenticationProvider);
        }

        private static async Task<Application> GetAzureApp()
        {
            _accessToken = await GetAccessToken();
            _graphClient = InitiateGraphServiceClient(_accessToken);
            return await _graphClient.Applications[ObjectId].Request().GetAsync();
        }

        public static async Task ExportUsersAssignedToApp()
        {
            var assignedUsers = await GetUsersAssignedToApp();
            await ExportToExcel(assignedUsers);
        }

        public static async Task<List<UserAssignmentData>> GetUsersAssignedToApp()
        {
            var app = await GetAzureApp();
            var appRoles =  _graphClient.Applications[app.Id].Request().GetAsync().Result;
            
            var userAssignments = new List<UserAssignmentData>();
            try
            {
                var assignments = await _graphClient.ServicePrincipals[AzureServicePrinciple].AppRoleAssignedTo
                    .Request()
                    .Top(998)
                    .GetAsync();

                foreach (var assignment in assignments.CurrentPage)
                {
                    var user = await _graphClient.Users[assignment.PrincipalId.ToString()].Request().GetAsync();

                    var appRole =
                        appRoles.AppRoles.FirstOrDefault(x => x.Id.ToString().Equals(assignment.AppRoleId?.ToString()));
                    
                    userAssignments.Add(new UserAssignmentData
                    {
                        //UserId = assignment.PrincipalId.ToString(),
                        UserDisplayName = assignment.PrincipalDisplayName,
                        UserPrincipalName = user.UserPrincipalName,
                        //AppRoleId = assignment.AppRoleId?.ToString(),
                        AppRoleDisplayValue = appRole?.Value,
                        AppRoleDisplayName = appRole?.DisplayName
                    });
                }

                while (assignments.NextPageRequest != null)
                {
                    assignments = await assignments.NextPageRequest.GetAsync();
                    foreach (var assignment in assignments.CurrentPage)
                    {
                        var user = await _graphClient.Users[assignment.PrincipalId.ToString()].Request().GetAsync();
                        
                        var appRole =
                            appRoles.AppRoles.FirstOrDefault(x => x.Id.ToString().Equals(assignment.AppRoleId?.ToString()));
                        
                        userAssignments.Add(new UserAssignmentData
                        {
                           // UserId = assignment.PrincipalId.ToString(),
                            UserDisplayName = assignment.PrincipalDisplayName,
                            UserPrincipalName = user.UserPrincipalName,
                           // AppRoleId = assignment.AppRoleId?.ToString(),
                            AppRoleDisplayValue = appRole?.Value,
                            AppRoleDisplayName = appRole?.DisplayName
                        });
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        

            return userAssignments;
        }

        private static Task ExportToExcel(List<UserAssignmentData> data)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    // Add a worksheet
                    var worksheet = package.Workbook.Worksheets.Add("Assigned Users");

                    // Add headers
                   // worksheet.Cells["A1"].Value = "UserId";
                    worksheet.Cells["A1"].Value = "UserDisplayName";
                    worksheet.Cells["B1"].Value = "UserPrincipalName";
                   // worksheet.Cells["C1"].Value = "AppRoleId";
                    worksheet.Cells["C1"].Value = "AppRoleDisplayValue";
                    worksheet.Cells["D1"].Value = "AppRoleDisplayName";

                    // Add data
                    for (var i = 0; i < data.Count; i++)
                    {
                        //worksheet.Cells[i + 2, 1].Value = data[i].UserId;
                        worksheet.Cells[i + 2, 1].Value = data[i].UserDisplayName;
                        worksheet.Cells[i + 2, 2].Value = data[i].UserPrincipalName;
                       // worksheet.Cells[i + 2, 3].Value = data[i].AppRoleId;
                        worksheet.Cells[i + 2, 3].Value = data[i].AppRoleDisplayValue;
                        worksheet.Cells[i + 2, 4].Value = data[i].AppRoleDisplayName;
                    }

                    // Save the file
                    var fileInfo = new FileInfo("AssignedUsers_PROD.xlsx");
                    package.SaveAs(fileInfo);
                    Console.WriteLine($"Exported to {fileInfo.FullName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting to Excel: {ex.Message}");
            }

            return Task.CompletedTask;
        }
    }

    public class UserAssignmentData
    {
        public string? UserId { get; set; }
        public string? UserDisplayName { get; set; }
        public string? UserPrincipalName { get; set; }
        public string? AppRoleId { get; set; }
        public string? AppRoleDisplayName { get; set; }
        
        public string? AppRoleDisplayValue { get; set; }
    }
}
