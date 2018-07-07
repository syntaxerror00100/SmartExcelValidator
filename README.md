<p>A smart way to validate excel data using NPOI</p>
<p>&nbsp;</p>
<p><strong>Install using nuget package manager console</strong></p>
<pre class="language-csharp"><code>Install-Package SmartExcelValidator -Version 2.1.0</code></pre>
<p>&nbsp;</p>
<p>Create an object of&nbsp;<strong>ExcelValidator&nbsp;</strong>feed your excel's&nbsp;file stream</p>
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
<p>Execute&nbsp;the validation</p>
<pre class="language-csharp"><code>  var listOfEmployee = new List&lt;EmployeeModel&gt;(); //IF NO ERROR  AUTOMAPPED RESULT WILL BE LOADED HERE

            var result = validator.Validate&lt;EmployeeModel&gt;(out listOfEmployee, CostumValidation);

            // CHECK IF WITH ERROR; 
            // IF NO ERROR DO WHATEVER YOU WANT
            bool isWithError = result.WithError;</code></pre>
<p>To further understand how to use this awesome&nbsp;library please download the sample project.</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
