using ILakshya.Dal;
using ILakshya.Model;
using ILakshya.WebApi.Jwt;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Linq;

namespace ILakshya.WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WphAuthenticationController : ControllerBase
    {
        private readonly IAuthenticationRepository _wphAuthentication;
        private readonly ITokenManager _tokenManager; // used in token manager for token
        private readonly WebPocHubDbContext _dbContext;

        public WphAuthenticationController(IAuthenticationRepository wphAuthentication, ITokenManager tokenManager, WebPocHubDbContext dbContext)
        {
            _wphAuthentication = wphAuthentication;
            _tokenManager = tokenManager;
            _dbContext = dbContext;
        }

        [HttpPost("RegisterUser")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public ActionResult Create(User user)
        {
            try
            {
                if (user.RoleId == 1) // Admin
                {
                    // Admin registration
                    var passwordHash = BCrypt.Net.BCrypt.HashPassword(user.Password);
                    user.Password = passwordHash;
                    var result = _wphAuthentication.RegisterUser(user);
                    if (result > 0)
                    {
                        return Ok();
                    }
                }
                else if (user.RoleId == 2) // Student
                {
                    // Student registration with a single password for all enrollments
                    var passwordHash = BCrypt.Net.BCrypt.HashPassword("Ganesh56"); // Set a default password for students
                    var studentEnrollNos = _dbContext.Students.Select(s => s.EnrollNo).ToList();
                    if (studentEnrollNos.Count == 0)
                    {
                        return BadRequest("No student enrollments found.");
                    }
                    foreach (var enrollNo in studentEnrollNos)
                    {
                        var studentUser = new User
                        {
                            EnrollNo = enrollNo.ToString(),
                            Password = passwordHash,
                            RoleId = 2 // Student role
                        };
                        _wphAuthentication.RegisterUser(studentUser);
                    }
                    return Ok();
                }
                else
                {
                    return BadRequest("Invalid RoleId.");
                }
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
            return BadRequest();
        }
        [HttpPost("CheckCredentials")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public ActionResult<AuthResponses> GetDetails(User user)
        {
            try
            {
                var authUser = _wphAuthentication.CheckCredentials(user);
                if (authUser == null)
                {
                    return NotFound();
                }
                if (!BCrypt.Net.BCrypt.Verify(user.Password, authUser.Password))
                {
                    return BadRequest("Incorrect Password! Please Check your Password");
                }
                var roleName = _wphAuthentication.GetUserRole(authUser.RoleId);

                var authResponse = new AuthResponses()
                {
                    IsAuthenticated = true,
                    Role = roleName,
                    Token = _tokenManager.GenerateToken(authUser, roleName),
                    EnrollNo = authUser.EnrollNo
                };
                return Ok(authResponse);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}

