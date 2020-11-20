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
    $Query = "EVALUATE ALL(AGG_Sales[Amount], AGG_Sales[Location])" 

    #Use XMLA to get data. Store it in $Results
    $Connection = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection
    $Results = New-Object System.Data.DataTable
    $Connection.ConnectionString = "Datasource=asazure://westeurope.asazure.windows.net/srvanalysis;initial catalog=DemoProject;User ID=app:c7c1b553-e6d5-4e08-9ae7-2ee96dc179c7@72f988bf-86f1-41af-91ab-2d7cd011db47;Password=N.FC4sNM7..NXVmf3Z~G5-Vtmz9PSIGl1Z" 
    $Connection.Open()
    $Adapter = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter $Query ,$Connection
    $Adapter.Fill($Results)
    $Connection.Dispose()
    $Connection.Close()

    #Connect to Database with SQL User
    $Database   = 'writeback'
    $Server     = 'moderndwh.database.windows.net'
    $UserName   = 'insert_user'
    $Password   = 'Swissre123456'
    $connectionString = 'Data Source={0};database={1};User ID={2};Password={3}' -f $Server,$Database,$UserName,$Password

    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $connectionString
    $sqlConnection.Open()

    #Create a bulkinsert object and insert
    $bc = new-object ("System.Data.SqlClient.SqlBulkCopy") $sqlConnection
    $bc.DestinationTableName = "dbo.AGG_Sales"
    $bc.WriteToServer($Results)

    #close SQL connection
    $sqlConnection.Close()
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
