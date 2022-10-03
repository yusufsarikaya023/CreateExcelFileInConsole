

ExcelWrapper wrapper = new ExcelWrapper();
wrapper.AddColumn("Name");
wrapper.AddColumn("Age");
wrapper.AddColumn("Address");
wrapper.AddColumn("Data", new List<string> { "1", "2", "3", "4" });
wrapper.CreateExcelFile("test");