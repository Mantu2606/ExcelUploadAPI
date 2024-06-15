using ILakshya.Dal;
using ILakshya.Model;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ILakshya.WebApi.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    public class TeacherController : ControllerBase
    {
        private readonly ICommonRepository<Teacher> _teacherRepository;
        public TeacherController(ICommonRepository<Teacher> repository)
        {
            _teacherRepository = repository;
        }
        [HttpGet]  // Get ALl API 
                   // [Authorize(Roles = "Employee, Hr")]
        public IEnumerable<Teacher> Get()
        {
            return _teacherRepository.GetAll();
        }
        [HttpGet("{id:int}")] // Get by Id 

        //[Authorize(Roles = "Employee, Hr")]
        public Teacher GetDetails(int id)
        {
            return _teacherRepository.GetDetails(id);
        }

        [HttpPost]
        [ProducesResponseType(StatusCodes.Status201Created)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        // [Authorize(Roles = "Employee, Hr")]
        public ActionResult Create(Teacher teacher)
        {
            _teacherRepository.Insert(teacher);
            var result = _teacherRepository.SaveChanges();
            if (result > 0)
            {
                return CreatedAtAction("GetDetails", new { id = teacher.TeacherId }, teacher);
            }
            return BadRequest();
        }

        [HttpPut]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]

        public ActionResult Update(Teacher teacher)
        {
            _teacherRepository.Update(teacher);
            var result = _teacherRepository.SaveChanges();
            if (result > 0)
            {
                return NoContent();
            }
            return BadRequest();
        }

        [HttpDelete("{id}")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [Authorize(Roles = "Hr")]
        public ActionResult<Teacher> Delete(int id)
        {
            var student = _teacherRepository.GetDetails(id);
            if (student == null)
            {
                return NotFound();
            }
            else
            {
                _teacherRepository.Delete(student);
                _teacherRepository.SaveChanges();
                return NoContent();
            }
        }
    }
}
