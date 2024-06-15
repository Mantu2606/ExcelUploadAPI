using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILakshya.Model
{
    public class User
    {
        public int UserId { get; set; }

        [MaxLength(100)]
        [Required(ErrorMessage = "Email Id is a required fields")]
        public string Email { get; set; } = string.Empty;

       /* [MaxLength(20)]
        [Required(ErrorMessage = "Enrollment number is required")]*/
        public string? EnrollNo { get; set; } = string.Empty;
        [MaxLength(100)]
        [Required(ErrorMessage = "Password is a required fileds")]

        public string Password { get; set; } = string.Empty;

        public int RoleId { get; set; } // 1-Admin, 2-Teacher, 3-Student
        public Role? Role { get; set; }
    }
}
