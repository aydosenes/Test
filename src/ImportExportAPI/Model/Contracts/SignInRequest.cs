using System;
using System.ComponentModel.DataAnnotations;

namespace ImportExportAPI.Model.Contracts
{
    public class SignInRequest
    {
        
        public string Name { get; set; }

        public string Surname { get; set; }

        public string Phone { get; set; }

        public string Email { get; set; }

    }
}
