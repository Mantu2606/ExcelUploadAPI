using ILakshya.Model;

namespace ILakshya.Dal
{
    public interface IAuthenticationRepository
    {
        int RegisterUser(User user);
        User? CheckCredentials(User user);
        string GetUserRole(int roleId);
    }
}
