using Xunit.Abstractions;
namespace BulkInsert;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // XSSF ve HSSF arasında seçim yapabilirsiniz, XSSF .xlsx dosyaları için, HSSF .xls dosyaları için.

public class UnitTest1
{
    private readonly ITestOutputHelper _testOutputHelper;

    public UnitTest1(ITestOutputHelper testOutputHelper)
    {
        _testOutputHelper = testOutputHelper;
    }

    public class FoodMenuData
    {
        public int prmUniversiteId { get; set; }
        public int prmFoodMenuForId { get; set; }
        public int prmFoodMenuTypeId { get; set; }
        public int prmFoodMenuTitleId { get; set; }
        public DateTime mealDate { get; set; }
        public bool isHoliday { get; set; }
        public string meal { get; set; }
        public int totalCalories { get; set; }
        public bool isActive { get; set; }
    }
    
    [Fact]
    public void Test1()
    {
       
        string debugFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

        string excelFilePath = Path.Combine(debugFolder, "DiningMenu_v1.xlsx");
        
        List<FoodMenuData> menuList = new List<FoodMenuData>();

        // Excel dosyasını oku
        using (FileStream fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
        {
            // Dosyayı bir XSSFWorkbook nesnesine yükle
            XSSFWorkbook workbook = new XSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(7); // Backend shettini al

            // Başlıkları atla, verileri al
            var menuData = new FoodMenuData();
            for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue; // Satır boşsa atla
                
                // Satırda herhangi bir hücre değeri var mı diye kontrol et
                bool hasCellValue = false;
                for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                {
                    ICell cell = row.GetCell(cellIndex);
                    if (cell != null && !string.IsNullOrEmpty(cell.ToString()) && cell.NumericCellValue>0)
                    {
                        hasCellValue = true;
                        break;
                    }
                }
                
                // Eğer satırda herhangi bir hücre değeri yoksa, bu satırı atla
                if (!hasCellValue) continue;
                try
                {
                     menuData = new FoodMenuData
                    {
                        prmUniversiteId =Convert.ToInt32(row.GetCell(0).NumericCellValue),
                        prmFoodMenuForId = Convert.ToInt32(row.GetCell(1).NumericCellValue),
                        prmFoodMenuTypeId =  Convert.ToInt32(row.GetCell(2).NumericCellValue),
                        prmFoodMenuTitleId = Convert.ToInt32(row.GetCell(3).NumericCellValue),
                        mealDate = row.GetCell(4).DateCellValue,
                        isHoliday = row.GetCell(5).BooleanCellValue,
                        meal = row.GetCell(6).StringCellValue,
                        totalCalories = Convert.ToInt32(row.GetCell(7).NumericCellValue),
                        isActive = row.GetCell(8).BooleanCellValue
                    };
                }
                catch (Exception ex)
                {
                    _testOutputHelper.WriteLine($"Hata: Satır {rowIndex + 1} - {ex.Message}");
                    throw;
                }
                menuList.Add(menuData);
            }
        }


        foreach (var item in menuList)
        {
            Console.WriteLine($"{item.prmUniversiteId} - {item.mealDate} - {item.meal}");
        }
    }
}