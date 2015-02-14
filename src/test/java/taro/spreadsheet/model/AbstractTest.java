package taro.spreadsheet.model;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Date;

/**
 * Created by David Gautier on 2/2/2015.
 */
public abstract class AbstractTest {

    protected SpreadsheetTab getSpreadsheetTab() {
        SpreadsheetWorkbook workbook = getSpreadsheetWorkbook();
        return new SpreadsheetTab(workbook, "testing");
    }

    private Object[][] mockSpreadsheetData = {
            //            A            B        C        D            E
        /* 1 */    {    "Fred",        7d,        null,    2.7d,        null        },
        /* 2 */    {    null,        -1d,    null,    4.3d,        new Date()    },
        /* 3 */    {    "Sam",        20d,    "OH",    .154d,        new Date()    },
        /* 4 */    {    "Mary",        null,    null,    -589.1d,    new Date()    },
        /* 5 */    {    "",            4d,        "B",    null,        new Date()    }
    };

    protected SpreadsheetTab getSpreadsheetTabWithValues(){
        SpreadsheetTab spreadsheetTab = getSpreadsheetTab();
        spreadsheetTab.setValue("A1","Fred");
        spreadsheetTab.setValue("B1",7d);
        spreadsheetTab.setValue("C1",null);
        spreadsheetTab.setValue("D1",2.7d);
        spreadsheetTab.setValue("E1",null);

        spreadsheetTab.setValue("A2",null);
        spreadsheetTab.setValue("B2",-1d);
        spreadsheetTab.setValue("C2",null);
        spreadsheetTab.setValue("D2",4.3d);
        spreadsheetTab.setValue("E2",new Date());

        spreadsheetTab.setValue("A3","Sam");
        spreadsheetTab.setValue("B3",20d);
        spreadsheetTab.setValue("C3","OH");
        spreadsheetTab.setValue("D3",.154d);
        spreadsheetTab.setValue("E3",new Date());

        spreadsheetTab.setValue("A4","Mary");
        spreadsheetTab.setValue("B4",null);
        spreadsheetTab.setValue("C4",null);
        spreadsheetTab.setValue("D4",-589.1d);
        spreadsheetTab.setValue("E4",new Date());

        spreadsheetTab.setValue("A5","");
        spreadsheetTab.setValue("B5",4d);
        spreadsheetTab.setValue("C5","B");
        spreadsheetTab.setValue("D5",null);
        spreadsheetTab.setValue("E5",new Date());

        return spreadsheetTab;
    }

    protected SpreadsheetWorkbook getSpreadsheetWorkbook() {
        XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        return new SpreadsheetWorkbook(poiWorkbook);
    }

    protected SpreadsheetCell getCell() {
        return getSpreadsheetTab().getOrCreateCell("A1");
    }

}
