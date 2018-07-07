# SmartExcelValidator
<p>A smart way to validate excel data using NPOI</p>
<p>&nbsp;</p>
<p><strong>Install using nuget package manager console</strong></p>
<pre class="language-csharp"><code>Install-Package SmartExcelValidator -Version 2.1.0</code></pre>
<p>&nbsp;</p>
<p>Create object of&nbsp;<strong>ExcelValidator&nbsp;</strong>feed your excel excel's&nbsp;filestream</p>

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
<p>The settings above code will start reading the column at Column <strong>B&nbsp;</strong>row <strong>2&nbsp;</strong>or (B2) and the actual data is at row 3.</p>
<p><strong>&nbsp;</strong></p>
<p><strong>&nbsp;</strong></p>
<table class="table table-bordered">
<tbody>
<tr>
<td>Test</td>
<td>asdf</td>
</tr>
<tr>
<td>Here</td>
<td>Now</td>
</tr>
</tbody>
</table>
<p><strong>&nbsp;</strong></p>
