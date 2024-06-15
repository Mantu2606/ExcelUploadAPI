using ILakshya.Model;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Data;
using System.Reflection.Emit;

namespace ILakshya.Dal
{
    public class WebPocHubDbContext : DbContext
    {
        public WebPocHubDbContext()
        {

        }
        public WebPocHubDbContext(DbContextOptions<WebPocHubDbContext> options) : base(options)
        {

        }
        protected override void OnConfiguring(DbContextOptionsBuilder options)
        {
            if (!options.IsConfigured)
            {
                options.UseSqlServer("Data Source = localhost; Initial Catalog = IdAdmin; Integrated Security = true; TrustServerCertificate=True");
            }
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Student>()
             .Property(s => s.EnrollNo)
             .IsRequired();

            modelBuilder.Entity<Teacher>().HasData(
                new Teacher() { TeacherId = 201, TeacherName = "Mohit Agrawal", Address = "Gorakhpur", Subject = "Hindi", Email = "Mohit@gmail.com", Phone = "+6388978442", Zipcode = "420322", Country = "India", Avatar = "/images/john-mark.png" },
                new Teacher() { TeacherId = 202, TeacherName = "Rupesh Patel", Address = "Deoria", Subject = "Engish", Email = "Rupesh@gmail.com", Phone = "+9876543219", Zipcode = "421303", Country = "India", Avatar = "/images/john-mark.png" }
                );
            modelBuilder.Entity<Role>().HasData(
            new Role() { RoleId = 1, RoleName = "Admin", RoleDescription = "Admin can be access data both" },
            new Role() { RoleId = 2, RoleName = "Student", RoleDescription = "Student show our data" },
            new Role() { RoleId = 3, RoleName = "Teacher", RoleDescription = "Teacher can be update data" }
        );
            base.OnModelCreating(modelBuilder);
        }
        public DbSet<Student> Students { get; set; }
        public DbSet<Role> Roles { get; set; }
        public DbSet<User> Users { get; set; }
    }
}

