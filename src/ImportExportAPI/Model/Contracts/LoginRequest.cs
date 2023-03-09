using System;
namespace ImportExportAPI.Model.Contracts
{
    public class LoginRequest
    {
        public LoginRequest()
        {
        }
        public String UserName { get; set; }

        public String Password { get; set; }
    }
}
