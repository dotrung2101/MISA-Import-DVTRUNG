using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MISA.Import.Entities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using MySqlConnector;
using Dapper;
using System.Data;

namespace MISA.Import.Controllers
{
    [Route("api/v1/customers")]
    [ApiController]
    public class CustomerController : ControllerBase
    {

        IDbConnection dbConnection;

        string connectionString = ""
                + "Host=47.241.69.179;"
                + "Port=3306;"
                + "User Id=dev;"
                + "Password=12345678;"
                + "Database = MF824-Import-DVTRUNG;"
                + "convert zero datetime=True";

        [HttpPost("import")]
        public IActionResult Import(IFormFile formFile)
        {
            
            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest();
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest();
            }


            var customers = GetAllRecordFromFile(formFile);

            Validate(customers);

            return Ok(customers);
        }

        private void Validate(List<Customer> customers)
        {
            foreach(Customer customer in customers)
            {
                CheckDuplicateInDatabase(customer);
                CheckCustomerGroupExist(customer);
            }


            //Check duplicate CustomerCode in file
            for(int i = 0; i < customers.Count; i++)
            {
                for(int j = 0; j < customers.Count; j++)
                {
                    if(i != j)
                    {
                        if(customers[i].CustomerCode == customers[j].CustomerCode)
                        {
                            customers[i].Status += "Mã khách hàng trùng lặp trên file";
                            break;
                        }
                    }
                }
            }

            //Check duplicate PhoneNumber in file
            for (int i = 0; i < customers.Count; i++)
            {
                for (int j = 0; j < customers.Count; j++)
                {
                    if (i != j)
                    {
                        if (customers[i].PhoneNumber == customers[j].PhoneNumber)
                        {
                            customers[i].Status += "Số điện thoại trùng lặp trên file";
                            break;
                        }
                    }
                }
            }

            //Check duplicate Email in file
            for (int i = 0; i < customers.Count; i++)
            {
                for (int j = 0; j < customers.Count; j++)
                {
                    if (i != j)
                    {
                        if (customers[i].Email == customers[j].Email)
                        {
                            customers[i].Status += "Email trùng lặp trên file";
                            break;
                        }
                    }
                }
            }
        }

        private void CheckCustomerGroupExist(Customer customer)
        {
            string sqlCommand = "Proc_GetCustomerGroupByCustomerGroupName";

            DynamicParameters parameters = new DynamicParameters();

            parameters.Add("@Name", customer.CustomerGroupName);


            using(dbConnection = new MySqlConnection(connectionString))
            {
                var cg = dbConnection.QueryFirstOrDefault<CustomerGroup>(sqlCommand, param: parameters, commandType: CommandType.StoredProcedure);
                if(cg == null)
                {
                    customer.Status += "Nhóm khách hàng không tồn tại\n";
                }
                else
                {
                    customer.CustomerGroupId = cg.CustomerGroupId;
                }
            }

            
        }

        private void CheckDuplicateInDatabase(Customer customer)
        {
            if(CheckAttributeExistInDatabase("CustomerCode", customer.CustomerCode))
            {
                customer.Status += "Mã khách hàng đã tồn tại trên database\n";
            }
            if (CheckAttributeExistInDatabase("PhoneNumber", customer.PhoneNumber))
            {
                customer.Status += "Số điện thoại đã tồn tại trên database\n";
            }
            if (CheckAttributeExistInDatabase("Email", customer.Email))
            {
                customer.Status += "Email đã tồn tại trên database\n";
            }
        }

        private bool CheckAttributeExistInDatabase(string attributeName, string value)
        {
            string sqlCommand = $"Proc_Check{attributeName}Exist";

            DynamicParameters parameters = new DynamicParameters();

            parameters.Add($"@{attributeName}", value);

            var exist = true;

            using (dbConnection = new MySqlConnection(connectionString))
            {
                exist = dbConnection.QueryFirstOrDefault<bool>(sqlCommand, param: parameters, commandType: CommandType.StoredProcedure);
            }

            return exist;
        }

        private List<Customer> GetAllRecordFromFile(IFormFile formFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var customers = new List<Customer>();

            using (var stream = new MemoryStream())
            {
                formFile.CopyToAsync(stream);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    for (int row = 3; row < worksheet.Dimension.Rows; row++)
                    {
                        Customer c = new Customer(
                            new Guid(),
                            worksheet.Cells[row, 2].Value == null ? null : worksheet.Cells[row, 2].Value.ToString(),
                            worksheet.Cells[row, 1].Value == null ? null : worksheet.Cells[row, 1].Value.ToString(),
                            worksheet.Cells[row, 3].Value == null ? null : worksheet.Cells[row, 3].Value.ToString(),
                            null,
                            worksheet.Cells[row, 5].Value == null ? null : worksheet.Cells[row, 5].Value.ToString(),
                            worksheet.Cells[row, 7].Value == null ? null : worksheet.Cells[row, 7].Value.ToString(),
                            worksheet.Cells[row, 8].Value == null ? null : worksheet.Cells[row, 8].Value.ToString(),
                            worksheet.Cells[row, 9].Value == null ? null : worksheet.Cells[row, 9].Value.ToString(),
                            worksheet.Cells[row, 10].Value == null ? null : worksheet.Cells[row, 10].Value.ToString(),
                            worksheet.Cells[row, 11].Value == null ? null : worksheet.Cells[row, 11].Value.ToString(),
                            worksheet.Cells[row, 4].Value == null ? null : worksheet.Cells[row, 4].Value.ToString());
                        customers.Add(c);
                    }
                }
            }

            return customers;
        }
    }
}
