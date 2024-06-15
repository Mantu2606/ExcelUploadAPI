using ILakshya.Model;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text; // UTF 

namespace ILakshya.WebApi.Jwt
{
    public class TokenManager : ITokenManager
    { 
     private readonly SymmetricSecurityKey _key;
    public TokenManager(IConfiguration config)
    {
        // _key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(config["WebPocHubJWT:Secret"]));
        _key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(config["WebPocHubJWT:Secret"]));// appsetting.json by taken 'WebPocHubJWT:Secret'
    }
    public string GenerateToken(User user, string roleName ) //, string roleName)
    {
        var claims = new List<Claim>{
               new Claim(JwtRegisteredClaimNames.Name,user.Email),
               new Claim(ClaimTypes.Role, roleName),
            };
        var credentials = new SigningCredentials(_key, SecurityAlgorithms.HmacSha256Signature);
        var tokenDescriptor = new SecurityTokenDescriptor
        {

            Subject = new ClaimsIdentity(claims),
            Expires = DateTime.Now.AddMinutes(60),
            SigningCredentials = credentials,
        };
        var tokenHandler = new JwtSecurityTokenHandler();
        var token = tokenHandler.CreateToken(tokenDescriptor);
        return tokenHandler.WriteToken(token);
    }

        public string GenerateToken(User user)
        {
            throw new NotImplementedException();
        }

        //public string GenerateToken(User user)
        //{
        //    throw new NotImplementedException();
        //}
    }
}
