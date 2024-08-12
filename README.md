[![build library](https://github.com/mnieto/Kevull.EPPLus.MultiHeader/actions/workflows/build.yml/badge.svg)](https://github.com/mnieto/Kevull.EPPLus.MultiHeader/actions/workflows/build.yml)

# Kevull.EPPLus.MultiHeader
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

![image](https://github.com/user-attachments/assets/af7b7d4b-b4c2-4146-b8eb-75bfa1ac8b39)
