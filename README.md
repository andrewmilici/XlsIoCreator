# XlsIoCreator
A small utility to assist developers in converting any List<T> object to an Excel spreadsheet, using the Syncfusion XlsIO libraries.
## What?
A lot of a developer's job (or mine that is) is to implement reporting features within business applications. These reports usually come in the form of Excel spreadsheets. To implement this I would normally have to follow the standard (and repetitive) steps of:
  1. Obtain and collate data into a IEnumerable (or List)
  2. Loop through data and write to the spreadsheet, ensuring all columns are populated correctly.
  3. Pass the spreadsheet back to the user.
  
This utility attempts to lessen the burden on step 2, and automatically populate a spreadsheet using reflection with a simple ToXlsIo() method.

## How?
  
  1. Ensure you have a reference to the XlsIO Binaries - refer to (https://help.syncfusion.com/file-formats/xlsio/overview) for more information.
  2. Load the XlsIoCreator.cs into your project.
  3. Call ToXlsIo or ToXlsIoBuffer to generate
  NOTE: By default the utility will use the field name as the column header. use the [Display] attribute (see examples below) to set a custom header
  
  ## Example Usage
```csharp
//Displays attributes in use to control certain rendering:
   public class Customer
    {
        [Display(DisplayName ="Date of Birth")] //Set the column title to "Date of Birth"
        public DateTime? DateOfBirth { get; set; }
        public string EmailAddress { get; set; }
        public string FirstName { get; set; }
        public int Id { get; set; }
        public string LastName { get; set; }
        public string TelephoneNumber { get; set; }
        [Currency] //Display this value as a current, adding $ and comma within number.
        public decimal? Salary { get; set; }
        [DateFormat(DateFormat = DateFormatAttribute.DateFormatEnum.TimeOnly)] //Display the time portion of the DateTime
        public DateTime? CallTime { get; set; }
    }


```csharp
//Returns a spreadsheet in an MVC application, back to the client.

var customers = db.Customers.ToList(); //Generate a list of data you wish to fill the spreadsheet with. 

var fileContents = customers.ToXlsIoBuffer(); //Return a byte array with the spreadsheet

return File(fileContents, System.Net.Mime.MediaTypeNames.Application.Octet, "Customers.xlsx"); //Return the file to the client
```

```csharp
//Saves a spreadsheet to the local disk.

var customers = db.Customers.ToList(); //Generate a list of data you wish to fill the spreadsheet with. 

var filePath = @"C:\Temp\Customers.xlsx";
var fileContents = customers.ToXlsIo(filePath); //Saves the spreadsheet to C:\Temp\Customers.xlsx
System.Diagnostics.Process.Start(filePath); //Open the spreadsheet using the default handler, usually excel. One good way to tell the user the report has finished generating is to show them.

```

## To Do
* Convert to a NuGet package to make it more portable.
* Convert to .NET Core.
* Use local Date formats & currencies instead of hard coded dd/mm/yyyy and dollar symbols.
* Options to allow the user to specify default fonts and behaviours with formatting
