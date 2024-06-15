using ILakshya.Model;

namespace ILakshya.WebApi.Jwt
{
    public interface ITokenManager
    {
        string GenerateToken(User user, string roleName); 
    }
}
