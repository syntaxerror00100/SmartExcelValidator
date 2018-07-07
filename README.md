
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
<pre><code>public class EmployeeModel
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
         
    }
    </code></pre>
<p>&nbsp;</p>
<p>&nbsp;</p>
asdfasdf
