using namespace System.Net


# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)
# Load DLLs
Add-Type -Path "Microsoft.Identity.Client.dll" 
Add-Type -Path "Microsoft.AnalysisServices.AdomdClient.dll" 

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.
$name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body.Name
}

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

if ($name) {
    #Dax Query -> outcome needs to match SQL table
    $Query = "<yourDax>" 

    #Use XMLA to get data. Store it in $Results
    $Connection = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection
    $Results = New-Object System.Data.DataTable
    $Connection.ConnectionString = "Datasource=<Your Analysis service>;initial catalog=<Your model>;User ID=app:<App_ID>@<Tenant_ID>;Password=>Secret of Service Principal>" 
    $Connection.Open()
    $Adapter = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter $Query ,$Connection
    $Adapter.Fill($Results)
    $Connection.Dispose()
    $Connection.Close()

    #Connect to Database with SQL User
    $Database   = '<Target DB>'
    $Server     = '<Azure Server address>'
    $UserName   = '<SQL User with read/write access>'
    $Password   = '<PW of SQL USer>'
    $connectionString = 'Data Source={0};database={1};User ID={2};Password={3}' -f $Server,$Database,$UserName,$Password

    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $connectionString
    $sqlConnection.Open()

    #Create a bulkinsert object and insert
    $bc = new-object ("System.Data.SqlClient.SqlBulkCopy") $sqlConnection
    $bc.DestinationTableName = "dbo.<Table to copy to>"
    $bc.WriteToServer($Results)

    #close SQL connection
    $sqlConnection.Close()
    $body = "Hello, $name. This HTTP triggered function executed successfully and read data from AAS and wrote it to SQL."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
