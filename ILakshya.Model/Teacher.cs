using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILakshya.Model
{
    public class Teacher
    {
        [Key]
        public int TeacherId { get; set; }

        [MaxLength(100)]
        public string TeacherName { get; set; } = string.Empty;

        [MaxLength(300)]
        public string Address { get; set; } = string.Empty;

        [MaxLength(50)]
        public string Subject { get; set; } = string.Empty;

        [MaxLength(10)]
        public string Zipcode { get; set; } = string.Empty;
        [MaxLength(50)]
        public string Country { get; set; } = string.Empty;

        [MaxLength(15)]
        [Required(ErrorMessage = "Teacher's phone number is a required field!")]
        public string Phone { get; set; } = string.Empty;

        [MaxLength(100)]
        [Required(ErrorMessage = "Teacher's email is a required field!")]

        [EmailAddress]
        public string Email { get; set; } = string.Empty;
        [MaxLength(200)]
        public string Avatar { get; set; } = string.Empty;

    }
}
