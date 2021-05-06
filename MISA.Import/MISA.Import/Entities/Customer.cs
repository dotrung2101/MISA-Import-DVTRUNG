using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MISA.Import.Entities
{
    public class Customer
    {
        public Customer(Guid id, string fullName, string customerCode, string memberCardCode, Guid? customerGroupId,
            string phoneNumber, string companyName, string taxCode, string email, string address, string note, string groupName, DateTime? dob)
        {
            CustomerId = id;
            FullName = fullName;
            CustomerCode = customerCode;
            MemberCardCode = memberCardCode;
            CustomerGroupId = customerGroupId;
            PhoneNumber = phoneNumber;
            CompanyName = companyName;
            TaxCode = taxCode;
            Email = email;
            Address = address;
            Note = note;
            CustomerGroupName = groupName;
            DateOfBirth = dob;
        }
        public Guid CustomerId { get; set; }
        public string FullName { get; set; }
        public string CustomerCode { get; set; }
        public string MemberCardCode { get; set; }
        public string CustomerGroupName { get; set; }
        public Guid? CustomerGroupId { get; set; }
        public string PhoneNumber { get; set; }
        public DateTime? DateOfBirth { get; set; }
        public string CompanyName { get; set; }
        public string TaxCode { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        public string Note { get; set; }
        public string Status { get; set; }
    }
}
