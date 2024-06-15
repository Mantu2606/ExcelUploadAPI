using ILakshya.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILakshya.Dal
{
    public interface ICommonRepository<T>
    {
        List<T> GetAll();
        T GetDetails(int id);
        void Insert(T item);
        void InsertRange(IEnumerable<T> items); // Add this line
        void Update(T item);
        void Delete(T item);
        int SaveChanges();
        Task<int> SaveChangesAsync();
    }
}
