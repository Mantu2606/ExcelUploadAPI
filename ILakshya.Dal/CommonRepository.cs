using ILakshya.Model;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILakshya.Dal
{
    public class CommonRepository<T> : ICommonRepository<T> where T : class
    {
        private readonly WebPocHubDbContext _dbContext;
        private DbSet<T> table;
        public CommonRepository(WebPocHubDbContext context)
        {
            _dbContext = context;
            table = _dbContext.Set<T>();
        }
        public List<T> GetAll()
        {
            return table.ToList();
        }
        public T GetDetails(int id)
        {
            return table.Find(id);
        }
        public void Insert(T item)
        {
            table.Add(item);
        }
        public void InsertRange(IEnumerable<T> items) // Add this method
        {
            table.AddRange(items);
        }
        public void Update(T item)
        {
            table.Attach(item);
            _dbContext.Entry(item).State = EntityState.Modified;
        }
        public void Delete(T item)
        {
            table.Remove(item);
        }
        public int SaveChanges()
        {
            return _dbContext.SaveChanges();
        }
        public async Task<int> SaveChangesAsync()
        {
            return await _dbContext.SaveChangesAsync();
        }

    }
}