
# SmartExcelValidator

<p>A smart way to validate excel data using NPOI</p>
<p>&nbsp;</p>
<ul>
<li><strong>1. Install using nuget package manager console</strong></li>
</ul>
<pre class="language-csharp"><code>Install-Package SmartExcelValidator -Version 2.1.0</code></pre>
<p>&nbsp;</p>
<ul>
<li><strong>Create a class model and use the data annotations</strong></li>
</ul>
<pre class="language-csharp"><code>    public class EmployeeModel
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
        [ExcelColumnName("Birth Date")]
        public DateTime BirthDate { get; set; }

        [Required]
        [ExcelColumnName("Department Name")]
        [TableDependency("departmentDT","Name","ID")]
        public int DepartmentId { get; set; }


        [TableDependency("locationDT", "Name", "ID")]
        [ExcelColumnName("Location Name")]
        public int LocationId { get; set; }
         
    }</code></pre>
<p>&nbsp;</p>
<p>&nbsp;</p>
<ul>
<li><strong>Create an object of&nbsp;ExcelValidator&nbsp;feed your excel's&nbsp;file stream</strong></li>
</ul>
<pre class="language-csharp"><code>
           var excelFileStream = new FileStream(@"EXCEL FILE PATH", FileMode.Open);


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
            };</code></pre>
<p>The settings above code will start reading the column at Column <strong>B&nbsp;</strong>row <strong>2&nbsp;</strong>or (B2) and the actual data is in row 3.</p>
<p><strong>&nbsp;</strong></p>
<ul>
<li><strong>Execute&nbsp;the validation, the validate method has an out parameter that will return a list of your object if there no error encountered in validations</strong></li>
</ul>
<pre class="language-csharp"><code>  var listOfEmployee = new List&lt;EmployeeModel&gt;(); //IF NO ERROR  AUTOMAPPED RESULT WILL BE LOADED HERE

            var result = validator.Validate&lt;EmployeeModel&gt;(out listOfEmployee, CostumValidation);

            // CHECK IF WITH ERROR; 
            // IF NO ERROR DO WHATEVER YOU WANT
            bool isWithError = result.IsWithError;</code></pre>
<p>To further understand how to use this awesome&nbsp;library please download the sample project.</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
