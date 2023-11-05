[![build library](https://github.com/mnieto/EPPlus.MultiHeader/actions/workflows/build.yml/badge.svg)](https://github.com/mnieto/EPPlus.MultiHeader/actions/workflows/build.yml)

# EPPlus.MultiHeader
Extension for the [EPPlus](https://github.com/EPPlusSoftware/EPPlus) library to create reports from complex objects

Given a list like this:
```csharp
            var complexObject = new List<RootLevel> { 
                new RootLevel {
                    SimpleProperty = "String1",
                    ComplexProperty = new SecondLevel
                    {
                        LeftColumn = "Left side 1",
                        RightColumn = new ThirdLevel
                        {
                            CatA = 11,
                            CatB = 12,
                            CatC = 13
                        }
                    }
                }, 
                new RootLevel {
                    SimpleProperty = "String2",
                    ComplexProperty = new SecondLevel
                    {
                        LeftColumn = "Left side 2",
                        RightColumn = new ThirdLevel
                        {
                            CatA = 21,
                            CatB = 22,
                            CatC = 23
                        }
                    }
                }
            };
```

this code:
```csharp
            using var xls = new ExcelPackage();
            var report = new MultiHeaderReport<RootLevel>(xls, "Object");
            report.GenerateReport(complexObject);
            xls.SaveAs("Report.xlsx");
```

will render like this:
![image](https://github.com/mnieto/EPPlus.MultiHeader/assets/7962206/4cb6e383-5d22-46f5-9308-ad21f82100ae)


