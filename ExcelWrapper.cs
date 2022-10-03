using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;

public class ExcelWrapper
{
    public string DateTimeWithFormat = DateTime.Now.ToString("ddMMyyyyHHmmss");

    public string[] sheetName = new string[] { "Product", "Data" };
    public ICollection<string> Headers { get; set; } = new HashSet<string>();
    public IDictionary<string, string[]> Formulas { get; set; } = new Dictionary<string, string[]>();

    public void CreateExcelFile(string fileName)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        var worksheetProduct = package.Workbook.Worksheets.Add(sheetName[0]);
        var worksheetData = package.Workbook.Worksheets.Add(sheetName[1]);


        for (int i = 1; i < Headers.Count + 1; i++)
        {
            var currentCell = worksheetProduct.Cells[1, i];
            var currentValue = Headers.ElementAt(i - 1);
            currentCell.Value = currentValue;
            var formula = Formulas.ContainsKey(Headers.ElementAt(i - 1)) ? Formulas[currentValue] : null;
            if (formula != null)
            {
                var index = Array.IndexOf(Formulas.Keys.ToArray(), currentValue);
                worksheetData.Cells[1, index + 1].Value = currentValue;
                worksheetData.Cells[2, index + 1].LoadFromCollection(formula);

                IExcelDataValidationList validation = worksheetProduct
                    .DataValidations
                    .AddListValidation(worksheetData.Cells[2, i, formula.Count() - 1 + i, i].Address);
                validation.ShowErrorMessage = true;
                validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                validation.ErrorTitle = "An invalid value was entered";
                validation.Error = "Select a value from the list";

                foreach (string opt in formula) validation.Formula.Values.Add(opt);

                validation.AllowBlank = true;
                validation.Validate();
            }

        }

        var file = new FileInfo(fileName);
        package.SaveAs($"{fileName}_{DateTimeWithFormat}.xlsx");
    }

    public void AddColumn(string header)
    {
        Headers.Add(header);
    }

    public void AddColumn(string header, List<string> formula)
    {
        Headers.Add(header);
        Formulas.Add(header, formula.ToArray());
    }
}