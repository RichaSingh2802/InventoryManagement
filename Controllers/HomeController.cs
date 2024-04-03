using ExcelDataReader;
using InventoryManagement.Models;
using InventoryManagement.Repository;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace InventoryManagement.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ApplicationDbContext _dbContext;

        /// <summary>
        /// Constructor to initialize context
        /// </summary>
        /// <param name="logger"></param>
        /// <param name="dbContext"></param>
        public HomeController(ILogger<HomeController> logger, ApplicationDbContext dbContext)
        {
            _logger = logger;
            _dbContext = dbContext;
        }

        /// <summary>
        /// Method to Dispaly the Upload Excel View
        /// </summary>
        /// <returns></returns>
        public IActionResult ImportInventory()
        {
            return View();
        }

        /// <summary>
        /// Post method for File Uplaod
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> ImportInventory(IFormFile file)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (file == null || file.Length <= 0)
                {
                    TempData["ErrorMessage"] = "No file selected.";
                    return RedirectToAction("ImportInventory");
                }

                if (!Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    TempData["ErrorMessage"] = "Invalid file format. Only Excel files (.xlsx) are supported.";
                    return RedirectToAction("ImportInventory");
                }

                if (file != null && file.Length > 0)
                {

                    var uploadsFolder = $"{Directory.GetCurrentDirectory()}\\wwwroot\\Uploads\\";

                    if (!Directory.Exists(uploadsFolder))
                    {
                        Directory.CreateDirectory(uploadsFolder);

                    }
                    var filePath = Path.Combine(uploadsFolder, file.FileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }
                    using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            bool isHeaderSkipped = false;

                            {
                                while (reader.Read())
                                {
                                    if (!isHeaderSkipped)
                                    {
                                        isHeaderSkipped = true;
                                        continue;
                                    }
                                    Product product = new Product();
                                    product.Name = reader.GetValue(0).ToString();
                                    product.Description = reader.GetValue(1).ToString();
                                    product.Quantity = Convert.ToInt32(reader.GetValue(2));
                                    product.Price = Convert.ToDecimal(reader.GetValue(3));
                                    product.CreatedDate=Convert.ToDateTime(reader.GetValue(4));
                                    product.UpdatedDate=Convert.ToDateTime(reader.GetValue(5));

                                    _dbContext.Add(product);
                                    await _dbContext.SaveChangesAsync();
                                }
                            } while (reader.NextResult()) ;

                        }
                    }

                }
                TempData["SuccessMessage"] = "File uploaded successfully.";
                return RedirectToAction("ImportInventory");
            }
            catch(Exception ex)
            {
                TempData["ErrorMessage"] = $"Error: {ex.Message}";
                return RedirectToAction("ImportInventory");
            }
            
        }
                       
    }
}
