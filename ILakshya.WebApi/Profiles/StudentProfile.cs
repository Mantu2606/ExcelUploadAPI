using AutoMapper;
using ILakshya.Model;
using ILakshya.WebApi.DTO;

namespace ILakshya.WebApi.Profiles
{
    public class StudentProfile : Profile
    {
        public StudentProfile()
        {
            CreateMap<Student, StudentDto>();
            CreateMap<NewStudentDto, Student>();
            CreateMap<UpdateStudentDto, Student>();
        }
    }
}
