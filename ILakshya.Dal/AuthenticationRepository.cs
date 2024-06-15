
 using ILakshya.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILakshya.Dal
{
    public class AuthenticationRepository : IAuthenticationRepository
    {

        private readonly WebPocHubDbContext _dbContext;
        public AuthenticationRepository(WebPocHubDbContext context)
        {
            _dbContext = context;
        }

        public User? CheckCredentials(User user)
        {
            var users = _dbContext.Users
                .Where(u => u.Email == user.Email || u.EnrollNo == user.EnrollNo)
                .ToList();


            if (users.Count > 1)
            {
                throw new InvalidOperationException("Multiple users found with the same Email or EnrollNo.");
            }
            return users.SingleOrDefault();
        }
        public string GetUserRole(int roleId)
        {
            return _dbContext.Roles.SingleOrDefault(role => role.RoleId == roleId).RoleName;
        }

        public int RegisterUser(User user)
        {
            var existingUser = _dbContext.Users
                .Any(u => u.Email == user.Email || u.EnrollNo == user.EnrollNo);

            if (existingUser)
            {
                throw new InvalidOperationException("A user with the same Email or EnrollNo already exists.");
            }
            _dbContext.Users.Add(user);
            return _dbContext.SaveChanges();
        }
    }
}

