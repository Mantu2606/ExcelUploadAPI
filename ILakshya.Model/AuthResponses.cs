using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILakshya.Model
{
    public class AuthResponses
    {
        public string Token { get; set; } = string.Empty;
        public string Role {  get; set; } = string.Empty;
        public bool IsAuthenticated {  get; set; } = false;
        public string EnrollNo { get; set; } = string.Empty;
      
    }
}
 