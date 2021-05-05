using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MISA.Import.Entities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace MISA.Import.Controllers
{
    [Route("api/v1/customers")]
    [ApiController]
    public class CustomerController : ControllerBase
    {
        [HttpPost("import")]
        public IActionResult Import(IFormFile formFile)
        {
            if(formFile == null || formFile.Length <= 0)
            {
                return BadRequest();
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest();
            }

            var customers = new List<Customer>();

            using(var stream = new MemoryStream())
            {
                formFile.CopyToAsync(stream);

                using(var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    for(int row = 3; row < worksheet.Dimension.Rows; row++)
                    {
                        Customer c = new Customer(
                            new Guid(),
                            worksheet.Cells[row,2].Value.ToString(),
                            worksheet.Cells[row,1].Value.ToString(),
                            worksheet.Cells[row,3].Value.ToString(),
                            null,
                            worksheet.Cells[row, 5].Value.ToString(),
                            worksheet.Cells[row, 7].Value.ToString(),
                            worksheet.Cells[row, 8].Value.ToString(),
                            worksheet.Cells[row, 9].Value == null? "" : worksheet.Cells[row, 9].Value.ToString(),
                            worksheet.Cells[row, 10].Value == null ? "" : worksheet.Cells[row, 10].Value.ToString(),
                            worksheet.Cells[row, 11].Value == null ? "" : worksheet.Cells[row, 11].Value.ToString());
                        customers.Add(c);
                    }
                }
            }

            return Ok(customers);
        }
    }
}
