# poi-xly
A java annotation framework for exporting, validating and importing excel files using Apache POI

## Workbook

A workbook is the java representation of the Excel file.
It looks like this:


    @XLYWorkbook
    public class TestWorkbook {
    
        @XLYSheet(name = "Bananas", type = TestBananas.class, columns = {
                @XLYColumn(field = "agencies", headerTitle = "Agencies"),
                @XLYColumn(field = "creationDate", datePattern = "dd/MM/YYYY", headerTitle = "Creation date"),
                @XLYColumn(field = "quantity", headerTitle = "Quantity"),
                @XLYColumn(field = "revenue", headerTitle = "Revenue") })
        private List<TestBananas> bananas;
    
        @XLYSheet(name = "Scenario", type = TestScenario.class, columns = {
                @XLYColumn(field = "name", mandatory = true, headerTitle = "Name") })
        private List<TestScenario> scenarios;
    
        public List<TestBananas> getBananas() {
            return bananas;
        }
        public void setBananas(List<TestBananas> bananas) {
            this.bananas = bananas;
        }
        public List<TestScenario> getScenarios() {
            return scenarios;
        }
        public void setScenarios(List<TestScenario> scenarios) {
            this.scenarios = scenarios;
        }
    }

The result Excel file will contain:
- 2 sheets 'Bananas' and 'Scenario'
- the first sheet 'Bananas' will contain 4 columns: Agencies, Creation date, Quantity and Revenue
- the second sheet will only contain 1 column: Name

## Export

To export java object to excel, use the following code:

    public static void main(String[] args) {
        final TestWorkbook workbook = new TestWorkbook();
        workbook.setBananas(getBananas());
        // workbook.setScenarios(...);
        xlyExporter = new XLYExporter(workbook);
        OutputStream outputStream = new FileOutputStream("/tmp/my-file.xlsx");
        xlyExporter.export(outputStream);
    }

    public static List<TestBananas> getBananas() {
        final List<TestBananas> bananas = new ArrayList<>();
        final TestBananas firstTrip = getBananas("FR", "AL", "email1@corp1.com");
        bananas.add(firstTrip);
        final TestBananas returnTrip = getBananas("AL", "FR", "email2@corp2.com");
        bananas.add(returnTrip);
        return bananas;
    }

This will result in an excel file generated under `/tmp/my-file.xlsx` matching the content of the workbook.

## Validate

You can specify validation constraints at the workbook level.

For example:

    @XLYWorkbook
    public class TestWorkbook {
    
        @XLYSheet(name = "Bananas", type = TestBananas.class, rowValidator = TestFlownRowConstraint.class, columns = {
                @XLYColumn(field = "agencies", headerTitle = "Agencies"),
                @XLYColumn(field = "creationDate", mandatory = true, datePattern = "dd/MM/YYYY", headerTitle = "Creation date"),
                @XLYColumn(field = "quantity", headerTitle = "Quantity"),
                @XLYColumn(field = "revenue", headerTitle = "Revenue"),
                @XLYColumn(field = "email", pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,6}$", headerTitle = "Email"),
                @XLYColumn(field = "origin", cellValidator = TestOriginConstraint.class, headerTitle = "Origin"),
                @XLYColumn(field = "destination", headerTitle = "Destination") })
        private List<TestBananas> bananas;
    
        public List<TestBananas> getBananas() {
            return bananas;
        }
    
        public void setBananas(List<TestBananas> bananas) {
            this.bananas = bananas;
        }
    }

This example contains:
- bananas sheets has a custom rowValidator
- creationDate column is mandatory
- email column use a pattern validator
- origin column has a custom cellValidator

Once the validation constraint are define, you need to trigger the validation with the following code:

    XLYValidator  xlyValidator = new XLYValidator(new DefaultConstraintLocator());
    xlyValidator.setWorkbookClass(TestWorkbook.class);
    final InputStream inputStream = // path to the excel file
    final boolean isValid = xlyValidator.isValid(inputStream, outputStream);

If the workbook is invalid, a new excel file is written to the outputStream with error cells displayed in RED so that the user can fix them.
 
## Import

The import allow you to transform and excel file into a Workbook (a.k.a a POJO with List).
 
_note:_ Usually you should check that the content of the excel file is valid using XLYValidator before trying to import it (create the java object list).

    XLYImporter<TestWorkbook> xlyImporter = new XLYImporter<>();
    xlyImporter.setWorkbookClass(TestWorkbook.class);
    final InputStream inputStream = // path to the excel file
    final TestWorkbook workbook = xlyImporter.save(inputStream);
    List<TestBananas> bananas = workbook.getBananas();
