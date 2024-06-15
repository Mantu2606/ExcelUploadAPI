using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILakshya.Model
{
    public class Role
    {
        public int RoleId { get; set; }
        [MaxLength(30)]
        public string RoleName { get; set; } = string.Empty;
        [MaxLength(300)]
        public string RoleDescription { get; set; }  = string.Empty;

        public ICollection<User>? Users { get; set; }
    }
}
