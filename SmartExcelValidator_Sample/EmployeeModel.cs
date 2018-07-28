using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SmartExcelValidator.DataAnnotation;
using SmartExcelValidator.ValidationClass;

namespace SmartExcelValidator_Sample
{
    public class EmployeeModel
    {

        [PrimaryKey]
        [Required]
        [ValidationRegex("^[0-9]*$")]
        [ExcelColumnName("Employee ID")]
        public string EmployeeId { get; set; }

        [Required]
        [ValidationRegex("^[a-zA-Z]")]
        [ExcelColumnName("First Name")]
        public string FirstName { get; set; }

        [Required]
        [ValidationRegex("^[a-zA-Z]")]
        [ExcelColumnName("Last Name")]
        public string LastName { get; set; }
        
        [Required]
        [ExcelColumnName("Department Name")]
        [TableDependency("departmentDT","Name","ID")]
        public string DepartmentId { get; set; }


        [TableDependency("locationDT", "Name", "ID")]
        [ExcelColumnName("Location Name")]
        public string LocationId { get; set; }
         
    }
}
