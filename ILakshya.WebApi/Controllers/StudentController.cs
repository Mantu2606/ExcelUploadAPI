using AutoMapper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ILakshya.Dal;
using ILakshya.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;


using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenXmlCellType = DocumentFormat.OpenXml.Spreadsheet.CellType;
using NpoiCellType = NPOI.SS.UserModel.CellType;
using MathNet.Numerics.Distributions;


namespace ILakshya.WebApi.Controllers
{
    // i have a student enroll id i want to login for student enroll id and a single password for all student enroll Id when login then open student enroll profile plz generate code
    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {
        private readonly ICommonRepository<Student> _studentRepository;
        private readonly WebPocHubDbContext _dbContext;
        private readonly IMapper _mapper;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IAuthenticationRepository _wphAuthentication;

        public StudentController(WebPocHubDbContext dbContext, ICommonRepository<Student> repository, IMapper mapper, IWebHostEnvironment webHostEnvironment, IAuthenticationRepository wphAuthentication)
        {
            _dbContext = dbContext;
            _studentRepository = repository;
            _wphAuthentication = wphAuthentication;
            _mapper = mapper;
            _webHostEnvironment = webHostEnvironment;
        }



        /* [HttpPost("UploadExcel")]
         public async Task<IActionResult> UploadExcel(IFormFile file)
         {
             if (file == null || file.Length == 0)
             {
                 return BadRequest("No file uploaded.");
             }

             var students = new List<Student>();

             using (var stream = new MemoryStream())
             {
                 await file.CopyToAsync(stream);
                 stream.Position = 0;

                 using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, false))
                 {
                     WorkbookPart workbookPart = doc.WorkbookPart;
                     Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();
                     WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                     SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                     var headers = new List<string>();
                     bool isFirstRow = true;

                     var existingStudents = _dbContext.Students.ToDictionary(s => s.EnrollNo);
                     foreach (Row row in sheetData.Elements<Row>())
                     {
                         if (isFirstRow)
                         {
                             headers = row.Elements<Cell>().Select(cell => GetCellValue(doc, cell)).ToList();
                             isFirstRow = false;
                             continue;
                         }

                         var student = new Student();
                         var cells = row.Elements<Cell>().ToArray();
                         if (cells.Length < 14) continue;

                         student.EnrollNo = cells.Length > 0 ? ParseCellValue(cells[0], doc) : null;

                         if (student.EnrollNo != null && existingStudents.TryGetValue(student.EnrollNo, out var existingStudent))
                         {
                             student = existingStudent;
                         }

                         student.Name = cells.Length > 1 ? GetCellValue(doc, cells[1]) : "Unknown";
                         student.FatherName = cells.Length > 2 ? GetCellValue(doc, cells[2]) : "Unknown";
                         student.RollNo = cells.Length > 3 ? ParseCellValue(cells[3], doc).ToString() : null;
                         student.GenKnowledge = cells.Length > 4 ? ParseCellValue(cells[4], doc) ?? 0 : 0;
                         student.Science = cells.Length > 5 ? ParseCellValue(cells[5], doc) ?? 0 : 0;
                         student.EnglishI = cells.Length > 6 ? ParseCellValue(cells[6], doc) ?? 0 : 0;
                         student.EnglishII = cells.Length > 7 ? ParseCellValue(cells[7], doc) ?? 0 : 0;
                         student.HindiI = cells.Length > 8 ? ParseCellValue(cells[8], doc) ?? 0 : 0;
                         student.HindiII = cells.Length > 9 ? ParseCellValue(cells[9], doc) ?? 0 : 0;
                         student.Computer = cells.Length > 10 ? ParseCellValue(cells[10], doc) ?? 0 : 0;
                         student.Sanskrit = cells.Length > 11 ? ParseCellValue(cells[11], doc) ?? 0 : 0;
                         student.Mathematics = cells.Length > 12 ? ParseCellValue(cells[12], doc) ?? 0 : 0;
                         student.SocialStudies = cells.Length > 13 ? ParseCellValue(cells[13], doc) ?? 0 : 0;
                         student.MaxMarks = 5;
                         student.PassMarks = 2;
                         if (student.EnrollNo == null)
                             continue;
                         students.Add(student);
                     }
                 }
             }

             try
             {
                 _dbContext.Students.AddRange(students);
                 await _dbContext.SaveChangesAsync();
             }
             catch (Exception ex)
             {
                 return StatusCode(500, $"Internal server error: {ex.Message}");
             }

             return Ok(students);
         }

         private string GetCellValue(SpreadsheetDocument doc, Cell cell)
         {
             if (cell.CellValue == null) return null;

             string value = cell.CellValue.InnerText;

             if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
             {
                 var stringTable = doc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                 value = stringTable.ElementAt(int.Parse(value)).InnerText;
             }

             return value;
         }

         private int? ParseCellValue(Cell cell, SpreadsheetDocument doc)
         {
             var value = GetCellValue(doc, cell);
             if (int.TryParse(value, out var intValue))
             {
                 return intValue;
             }
             return null;
         }
 */
        [HttpPost("UploadExcel")]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            var students = new List<Student>();
            var logins = new List<User>(); // Sachine sir 
            try
            {
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    stream.Position = 0;

                    IWorkbook workbook = null ;
                    if (file.FileName.EndsWith(".xls"))
                    {
                        // For Excel 97-2003 format (XLS)
                        workbook = new HSSFWorkbook(stream);
                    }
                    else if (file.FileName.EndsWith(".xlsx"))
                    {
                        // For Excel 2007+ format (XLSX)
                        workbook = new XSSFWorkbook(stream);
                    }
                    if (workbook == null)
                    {
                        return BadRequest("Unsupported file"); 
                    }
                    var sheet = workbook.GetSheetAt(0); // Assuming only one sheet

                    var existingStudents = _dbContext.Students.ToDictionary(s => s.EnrollNo);

                    for (int rowIdx = 1; rowIdx <= sheet.LastRowNum; rowIdx++) // Start from 1 to skip header
                    {
                        var row = sheet.GetRow(rowIdx);
                        if (row == null) continue; // Skip empty rows

                        var student = new Student();

                        student.EnrollNo = ParseCellValue(row.GetCell(0));
                        if (student.EnrollNo != null && existingStudents.TryGetValue(student.EnrollNo, out var existingStudent))
                        {
                            student = existingStudent;
                        }

                        student.Name = GetCellValue(row.GetCell(1));
                        student.FatherName = GetCellValue(row.GetCell(2));
                        student.RollNo = ParseCellValue(row.GetCell(3))?.ToString();
                        student.GenKnowledge = ParseCellValue(row.GetCell(4)) ?? 0;
                        student.Science = ParseCellValue(row.GetCell(5)) ?? 0;
                        student.EnglishI = ParseCellValue(row.GetCell(6)) ?? 0;
                        student.EnglishII = ParseCellValue(row.GetCell(7)) ?? 0;
                        student.HindiI = ParseCellValue(row.GetCell(8)) ?? 0;
                        student.HindiII = ParseCellValue(row.GetCell(9)) ?? 0;
                        student.Computer = ParseCellValue(row.GetCell(10)) ?? 0;
                        student.Sanskrit = ParseCellValue(row.GetCell(11)) ?? 0;
                        student.Mathematics = ParseCellValue(row.GetCell(12)) ?? 0;
                        student.SocialStudies = ParseCellValue(row.GetCell(13)) ?? 0;
                        student.MaxMarks = 5; // Example values, adjust as needed
                        student.PassMarks = 2;
                        if (student.RollNo == null)
                            continue;
                        students.Add(student);

                        // Sachine sir 
                        logins.Add(new User() { 
                          Email = student.EnrollNo.ToString(),
                          EnrollNo = student.EnrollNo.ToString(),
                          Password = BCrypt.Net.BCrypt.HashPassword(student.EnrollNo+"_p@11"),
                          RoleId = 2
                    });
                    }
                }

                _dbContext.Students.AddRange(students); // Sachine sir 
                _dbContext.Users.AddRange(logins); // sachine sir 
                await _dbContext.SaveChangesAsync(); 
                return Ok(students);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        private string GetCellValue(NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null) return null;

            switch (cell.CellType)
            {
                case NpoiCellType.String:
                    return cell.StringCellValue;
                case NpoiCellType.Numeric:
                    if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue.ToString(); // Handle date values as needed
                    else
                        return cell.NumericCellValue.ToString();
                case NpoiCellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case NpoiCellType.Formula:
                    return cell.CellFormula; // Handle formula if needed
                default:
                    return null;
            }
        }

        private int? ParseCellValue(NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null || cell.CellType == NpoiCellType.Blank)
                return null;

            switch (cell.CellType)
            {
                case NpoiCellType.Numeric:
                    return (int)Math.Round(cell.NumericCellValue);
                case NpoiCellType.String:
                    if (int.TryParse(cell.StringCellValue, out int intValue))
                        return intValue;
                    return null;
                default:
                    return null;
            }
        }


        [HttpGet]
        public IEnumerable<Student> GetAll()
        {
            return _studentRepository.GetAll();
        }

        [HttpGet("{id:int}")]
        public ActionResult<Student> GetById(int id)
        {
            var student = _studentRepository.GetDetails(id);
            if (student == null)
            {
                return NotFound();
            }
            return Ok(student);
        }

        [HttpGet("ByEnrollNo/{enrollNo}")]
        public ActionResult<Student> GetStudentDetailsByEnrollNo(string enrollNo)
        {
            if (string.IsNullOrEmpty(enrollNo))
            {
                return BadRequest("EnrollNo cannot be null or empty.");
            }

            var student = _studentRepository.GetAll().FirstOrDefault(s => s.EnrollNo != null && s.EnrollNo.ToString() == enrollNo);
            if (student == null)
            {
                return NotFound("Student Not found");
            }

            return Ok(student);
        }

        [HttpPost("UploadProfilePicture/{id}")]
        public async Task<IActionResult> UploadProfilePicture(int id, IFormFile file)
        {
            var student = _studentRepository.GetDetails(id);
            if (student == null)
            {
                return NotFound("Student not found.");
            }

            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            // Save the file to the wwwroot/images folder
            var uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "images");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            var uniqueFileName = Guid.NewGuid().ToString() + "_" + file.FileName;
            var filePath = Path.Combine(uploadsFolder, uniqueFileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(fileStream);
            }

            // Update the profile picture path in the database
            student.ProfilePicture = "/images/" + uniqueFileName;
            _studentRepository.Update(student);

            

            

            _studentRepository.SaveChanges();

            return Ok("Profile picture uploaded successfully.");
        }

        [HttpDelete("{id}")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public ActionResult<Student> Delete(int id)
        {
            var student = _studentRepository.GetDetails(id);
            if (student == null) return NotFound();

            _studentRepository.Delete(student);
            _studentRepository.SaveChanges();
            return NoContent();
        }

        [HttpDelete("ByEnrollNo/{enrollNo}")]
        public ActionResult<Student> DeleteByEnrollNo(string enrollNo)
        {
            var student = _studentRepository.GetAll().FirstOrDefault(s => s.EnrollNo?.ToString() == enrollNo);
            if (student == null)
            {
                return NotFound();
            }
            _studentRepository.Delete(student);
            _studentRepository.SaveChanges();
            return NoContent();
        }

       
    }
}
