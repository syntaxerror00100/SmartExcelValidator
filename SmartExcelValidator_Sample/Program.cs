using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SmartExcelValidator;
using SmartExcelValidator.CommonClass;


namespace SmartExcelValidator_Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            // LOAD EXCEL FILE
            var excelFileStream = new FileStream(@"..\..\ExcelToValidateAndUpload.xlsx", FileMode.Open);


            var validator = new ExcelValidator(excelFileStream)
            {
                Settings =
                {
                    WorkSheetName                     = "Employee",
                    ColoumnStartAtColumnNumber        = 2,
                    ColoumnStartAtRowNumber           = 2,
                    DataStartAtRowNumber              = 3,
                    RowByRowColumnNumberLocation      = 1,
                    TableDependencyDataset            = GetTableDependencies(),
                    TableNameForPrimaryUniqueChecking = "employeesDT",
                    AutoMapType                       = Enums.AutomapType.AutoMapWithForienKeyDependency,
                    StatusErrorAnnotationMethod       = Enums.StatErrorAnnotationMethod.RowByRowInOneColumnWithStatus
                }
            };
            // COLUMN NUMBER WHERE TO PUT THE ERROR MESSAGES / ANNOTATIONS
            // SET THIS PROPERTY IF YOU HAVE TABLE DEPENDENCIES
            // I WANT EXCELVALIDATOR TO GET THE ID AUTOMATICALLY

            var listOfEmployee = new List<EmployeeModel>(); //IF NO ERROR  AUTOMAPPED RESULT WILL BE LOADED HERE

            var result = validator.UploadExcelAndDoTheDirtyJob<EmployeeModel>(out listOfEmployee, CostumValidation);

            // CHECK IF WITH ERROR; 
            // IF NO ERROR DO WHATEVER YOU WANT
            // ACCES THE LIS
            bool isWithError = result.WithError;

            File.WriteAllBytes(@"X:\test.xlsx",result.ExcelResultStream.ToArray());
            


        }


        /// <summary>
        /// Costums the validation.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="cellValue">The cell value.</param>
        /// <returns>First: error message, Second: new cell value if you want to replace</returns>
        static (string, string) CostumValidation(IRow row,string columnName,string cellValue)
        {
            // COSTUM VALIDATION FOR BIRTHDAY
            if (columnName == "BirthDate")
            {
                var isValidDate = DateTime.TryParse(cellValue,out var dateVal);

                if (isValidDate)
                    return ("", ""); // SET THE ITEM2 VALUE IF YOU WANT TO CHANGE THE VALUE

                return ($"Birth date is not a valid date", "");
            }



            return ("","");
        }

       static DataSet GetTableDependencies()
        {
            // WILL RETURN ALL TABLE DEPENDENCIES FOR DEPARTMENT AND LOCATION
            // FOR SAMPLE PORPUSES WILL JUST MOCKS THE DATA


            // DEPARTMENT
            var departmentDT = new DataTable("departmentDT");// THE DATATABLE NAME WILL BE USED IN TABLE DEPENDENCY ANNOATION
            departmentDT.Columns.Add("ID", typeof(int));
            departmentDT.Columns.Add("Name", typeof(string));
            
            departmentDT.Rows.Add(1, "HR");
            departmentDT.Rows.Add(1, "FINANCE");



            // LOCATION
            var locationDT = new DataTable("locationDT");// THE DATATABLE NAME WILL BE USED IN TABLE DEPENDENCY ANNOATION
            locationDT.Columns.Add("ID", typeof(int));
            locationDT.Columns.Add("Name", typeof(string));
            locationDT.Rows.Add(1, "MAKATI");
            locationDT.Rows.Add(2, "CEBU");



            // LIST OF ALL EMPLOYEES IN DATABASE
            var employeesDT = new DataTable("employeesDT");
            employeesDT.Columns.Add("ID", typeof(int));
            employeesDT.Columns.Add("FirstName", typeof(string));
            employeesDT.Columns.Add("LastName", typeof(string));
            employeesDT.Columns.Add("BirthDate", typeof(DateTime));
            employeesDT.Columns.Add("DepartmentId", typeof(int));
            employeesDT.Columns.Add("LocationId", typeof(int));

            employeesDT.Rows.Add(1, "Mark", "Lester", DateTime.Now, 1, 1);



            var result = new DataSet();
            result.Tables.Add(departmentDT);
            result.Tables.Add(locationDT);
            result.Tables.Add(employeesDT);



            return result;
        }



    }
}
