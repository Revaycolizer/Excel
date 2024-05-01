using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Globalization;
using WebApplication3.Models;
using WebApplication3.Services;

namespace WebApplication3.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ApplicationDbContext context;

        public HomeController(ILogger<HomeController> logger,ApplicationDbContext context)
        {
            _logger = logger;
            this.context = context;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult UploadExcel()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            if (file != null && file.Length > 0)
            {
                var uploadDirectory = $"{Directory.GetCurrentDirectory()}\\wwwroot\\UploadExcel";

                if (!Directory.Exists(uploadDirectory))
                {
                    Directory.CreateDirectory(uploadDirectory);
                }

                var filePath = Path.Combine(uploadDirectory, file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                

                    List<Student> studentsToInsert = new List<Student>();

                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        while (reader.Read())
                        {
                            string dateValue = reader.GetValue(2).ToString();
                            DateTime birthDate;

                            DateTime.TryParseExact(dateValue, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out birthDate);

                            string phoneNoValue = reader.GetValue(6).ToString();
                            long phoneNo;

                            long.TryParse(phoneNoValue, out phoneNo);

                            Student s = new Student
                            {
                                firstname = reader.GetValue(1).ToString(),
                                birthdate = birthDate.Date,
                                middlename = reader.GetValue(3).ToString(),
                                lastname = reader.GetValue(4).ToString(),
                                guardian = reader.GetValue(5).ToString(),
                                phoneno = phoneNo,
                                createdAt = DateTime.Now
                            };

                            studentsToInsert.Add(s);
                        }
                    }

                    context.Students.AddRange(studentsToInsert);
                    await context.SaveChangesAsync();

                    return RedirectToAction("Index"); // Assuming you want to redirect to the index page after inserting students

                }

            }

            return View();
        }
    }
}
