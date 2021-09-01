
package StingRaySharp.Parser.Utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

import StingRaySharp.Parser.StateUsAbriviationEnums;

public class XlsxParserUtility {

    final static ObjectMapper objectMapper = new ObjectMapper();

    public static void main( String[] args ) {

        try {
            String inFile = "C:\\Users\\noman\\Desktop\\StartLienSummaryChart.xlsx";
            String inFile2 = "C:\\Users\\noman\\Documents\\Varicent Start.xlsx";

            String outFile = "C:\\Users\\noman\\Desktop\\StartLienSummaryChart" + new Date().getMinutes() + "_" + new Date().getSeconds()
                    + ".xlsx";
            readXlsx( inFile2, outFile );

        } catch ( Exception e ) {
            // TODO: handle exception
        }

    }

    public static String readXlsx( String inFile, String outFile ) throws IOException {

        // String inFile = "C:\\Users\\noman\\Desktop\\StartLienSummaryChart.xlsx";

        FileInputStream fis = new FileInputStream( new File( inFile ) );

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook( fis );

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt( 0 );

        List< List< String > > myList = new ArrayList<>();

        // Get iterator to all the rows in current sheet
        Iterator< Row > rowIterator = mySheet.iterator();

        // Traversing over each row of XLSX file
        while ( rowIterator.hasNext() ) {
            Row row = rowIterator.next();

            List< String > innerList = new ArrayList<>();

            // For each row, iterate through each columns

            for ( int i = 0; i < 5; i++ ) {
                Cell cell = row.getCell( i );
                if ( !cell.toString().isEmpty() ) {
                    innerList.add( cell.toString() );
                    System.out.println( cell.toString() );
                }
            }

            /*
             Iterator< Cell > cellIterator = row.cellIterator(); 
               while ( cellIterator.hasNext() ) {
                Cell cell = cellIterator.next();
            }*/

            if ( !innerList.isEmpty() ) {
                myList.add( innerList );
            }
        }

        myWorkBook.close();

        myList.remove( 0 );
        myList.remove( 0 );

        System.out.println( "******************main list" );
        System.out.println( objectMapper.writeValueAsString( myList ) );

        List< Map< String, String > > rawDataList = new ArrayList<>();

        Iterator< List< String > > myListIterator = myList.iterator();
        while ( myListIterator.hasNext() ) {

            Map< String, String > rowInnerDataMap = new HashMap<>();

            List< String > innerDataList = myListIterator.next();

            String debtrNames = innerDataList.get( 0 );
            String jurisdiction = innerDataList.get( 1 );
            String service = innerDataList.get( 2 );
            String dateThru = innerDataList.get( 3 );
            String results = "";
            try {
                results = innerDataList.get( 4 );
            } catch ( Exception e ) {
            }
            String status = getStatusFromResults( results );

            /*      System.out.println( debtrNames );
            System.out.println( jurisdiction );
            System.out.println( dateThru );
            System.out.println( results );
            System.out.println( " " );*/

            String[] jurisdictionSplit = jurisdiction.split( "-" );
            if ( jurisdictionSplit != null && jurisdictionSplit.length > 0 ) {
                String usps = jurisdictionSplit[ 0 ];
                String localAreaName = jurisdictionSplit[ 1 ];
                String finalJurisdictionSplit = "";
                if ( localAreaName.contains( "Secretary" ) ) {

                    service = service.replace( "Certified", "" ).replace( "-", "" );

                    finalJurisdictionSplit = usps + " " + localAreaName + "\n" + "(" + service.trim() + ")";
                    // dateThru = dateThru + "\n" + "Certified";
                } else {

                    finalJurisdictionSplit = localAreaName + " " + usps + "\n" + "(" + service.trim() + ")";
                }

                rowInnerDataMap.put( "Debtor Name", debtrNames );
                rowInnerDataMap.put( "Status", status );
                rowInnerDataMap.put( "Jurisdiction", finalJurisdictionSplit );
                rowInnerDataMap.put( "State", usps );
                rowInnerDataMap.put( "Search Thru Date", dateThru );
            }

            rawDataList.add( rowInnerDataMap );
        }

        System.out.println( "******************rawDataList" );
        System.out.println( objectMapper.writeValueAsString( rawDataList ) );

        List< String > orderList = new ArrayList<>();
        Map< String, List< Map< String, String > > > finalRowDataMap = prepareFinalXlsxData( rawDataList, orderList );

        return createXlsxFile( finalRowDataMap, orderList, outFile );

    }

    private static String getStatusFromResults( String results ) {
        String status = "";
        if ( results != null && !results.isEmpty() ) {
            boolean resultFind = true;
            String lines[] = results.split( "\\r?\\n" );
            for ( String resultsLines : lines ) {
                if ( resultFind ) {
                    String[] resultSplit = resultsLines.split( " " );
                    if ( resultSplit != null && resultSplit.length > 0 ) {
                        if ( resultSplit[ 0 ].equals( "0" ) ) {

                        } else {
                            resultFind = false;
                            break;
                        }
                    }
                }
            }

            if ( resultFind ) {
                status = "Clear";
            } else {
                status = "Active";
            }
        } else {
            status = "";
        }
        return status;
    }

    private static Map< String, List< Map< String, String > > > prepareFinalXlsxData( List< Map< String, String > > rawDataList,
            List< String > orderList ) throws JsonProcessingException {

        List< Map< String, String > > preparedDebtorList = new ArrayList<>();
        getNextDebtorRecord( preparedDebtorList, rawDataList );

        System.out.println( "******************preparedDebtorList" );
        System.out.println( objectMapper.writeValueAsString( preparedDebtorList ) );

        int index = 1;
        Map< String, List< Map< String, String > > > finalRowDataMap = new HashMap<>();
        groupByState( preparedDebtorList, finalRowDataMap, index, orderList, null );

        System.out.println( "******************finalRowDataMap" );
        System.out.println( objectMapper.writeValueAsString( finalRowDataMap ) );

        return finalRowDataMap;
    }

    private static void groupByState( List< Map< String, String > > preparedDebtorList,
            Map< String, List< Map< String, String > > > finalRowDataMap, int index, List< String > orderList,
            String dbtMatchNameToResetIndex ) {

        String dbtrName = null;

        if ( !preparedDebtorList.isEmpty() ) {
            Map< String, String > innerDataMap = preparedDebtorList.get( 0 );
            String state = innerDataMap.get( "State" );
            dbtrName = innerDataMap.get( "Debtor Name" );

            if ( dbtMatchNameToResetIndex != null && !dbtMatchNameToResetIndex.equalsIgnoreCase( dbtrName ) ) {
                index = 1;
            }

            List< Map< String, String > > prepListForRowDataMap = new ArrayList<>();
            boolean match = false;
            Iterator< Map< String, String > > rawDataListIterator = preparedDebtorList.iterator();
            while ( rawDataListIterator.hasNext() ) {
                Map< String, String > preparedFinalRowsMap = new HashMap<>();

                Map< String, String > innerMap = rawDataListIterator.next();

                if ( innerMap.get( "State" ).equalsIgnoreCase( state ) && innerMap.get( "Debtor Name" ).equalsIgnoreCase( dbtrName ) ) {

                    String debtorName = innerMap.get( "Debtor Name" );
                    String Jurisdiction = innerMap.get( "Jurisdiction" );
                    String SearchThruDate = innerMap.get( "Search Thru Date" );
                    String status = innerMap.get( "Status" );

                    preparedFinalRowsMap.put( "#", String.valueOf( index ) );
                    preparedFinalRowsMap.put( "Debtor Name", debtorName );
                    preparedFinalRowsMap.put( "Jurisdiction", Jurisdiction );
                    preparedFinalRowsMap.put( "File No and Date", "" );
                    preparedFinalRowsMap.put( "Secured Party", "" );
                    preparedFinalRowsMap.put( "Collateral", "" );
                    preparedFinalRowsMap.put( "Search Thru Date", SearchThruDate );
                    preparedFinalRowsMap.put( "Status", status );

                    index++;
                    match = true;
                    prepListForRowDataMap.add( preparedFinalRowsMap );
                    rawDataListIterator.remove();
                }
            }
            if ( match ) {
                orderList.add( dbtrName + " " + "(" + StateUsAbriviationEnums.getKeyByValue( state ) + ")" );
                finalRowDataMap.put( dbtrName + " " + "(" + StateUsAbriviationEnums.getKeyByValue( state ) + ")", prepListForRowDataMap );
            }
        }

        if ( !preparedDebtorList.isEmpty() ) {
            groupByState( preparedDebtorList, finalRowDataMap, index, orderList, dbtrName );
        } else {
            return;
        }

    }

    private static void getNextDebtorRecord( List< Map< String, String > > preparedDebtorList, List< Map< String, String > > rawDataList ) {

        String searchNextDebtor = "";
        if ( !rawDataList.isEmpty() ) {
            Map< String, String > findDebtorMap = rawDataList.get( 0 );
            searchNextDebtor = findDebtorMap.get( "Debtor Name" );
            String state = findDebtorMap.get( "State" );

            Iterator< Map< String, String > > rawDataListIterator = rawDataList.iterator();
            while ( rawDataListIterator.hasNext() ) {
                Map< String, String > innerMap = rawDataListIterator.next();
                if ( innerMap.get( "Debtor Name" ).equalsIgnoreCase( searchNextDebtor )
                        && innerMap.get( "State" ).equalsIgnoreCase( state ) ) {
                    preparedDebtorList.add( innerMap );
                    rawDataListIterator.remove();
                }
            }
        }

        if ( !rawDataList.isEmpty() ) {
            getNextDebtorRecord( preparedDebtorList, rawDataList );
        } else {
            return;
        }

    }

    private static String createXlsxFile( Map< String, List< Map< String, String > > > finalRowDataMap, List< String > orderList,
            String outFile ) throws IOException {

        // String OutFileResult = "C:\\Users\\noman\\Desktop\\FIg\\" + new Date().getSeconds() + "_LienSummaryResul" + ".xlsx";

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet( "TelosLegalCorp" );

        List< String > headers = new ArrayList<>();
        headers.add( "#" );
        headers.add( "Debtor Name" );
        headers.add( "Jurisdiction" );
        headers.add( "File No and Date" );
        headers.add( "Secured Party" );
        headers.add( "Collateral" );
        headers.add( "Search Thru Date" );
        headers.add( "Status" );

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontName( "Arial" );
        headerStyle.setWrapText( true );
        headerStyle.setAlignment( HorizontalAlignment.CENTER );
        headerStyle.setVerticalAlignment( VerticalAlignment.CENTER );
        font.setBold( true );
        headerStyle.setFont( font );
        headerStyle.setBorderTop( BorderStyle.THIN );
        headerStyle.setBorderBottom( BorderStyle.THIN );
        headerStyle.setBorderLeft( BorderStyle.THIN );
        headerStyle.setBorderRight( BorderStyle.THIN );

        XSSFCellStyle titalStyle = workbook.createCellStyle();
        titalStyle.setFillForegroundColor( IndexedColors.GREY_25_PERCENT.getIndex() );
        titalStyle.setFillPattern( FillPatternType.SOLID_FOREGROUND );
        titalStyle.setFont( font );
        titalStyle.setWrapText( true );
        titalStyle.setAlignment( HorizontalAlignment.CENTER );
        titalStyle.setVerticalAlignment( VerticalAlignment.CENTER );

        XSSFCellStyle rowStyle = workbook.createCellStyle();
        rowStyle.setAlignment( HorizontalAlignment.CENTER );
        rowStyle.setVerticalAlignment( VerticalAlignment.CENTER );
        rowStyle.setWrapText( true );
        rowStyle.setBorderTop( BorderStyle.THIN );
        rowStyle.setBorderBottom( BorderStyle.THIN );
        rowStyle.setBorderLeft( BorderStyle.THIN );
        rowStyle.setBorderRight( BorderStyle.THIN );

        sheet.setDefaultRowHeightInPoints( 60 );

        Row header = sheet.createRow( 0 );
        for ( int i = 0; i < headers.size(); i++ ) {
            String headName = headers.get( i );
            Cell headerCell = header.createCell( i );
            headerCell.setCellValue( headName );
            headerCell.setCellStyle( headerStyle );
            // sheet.autoSizeColumn(i);
            // sheet.setColumnWidth(i,sheet.getColumnWidth(i)*17/10);
        }

        int rowInt = 0;

        for ( String rowOrderKey : orderList ) {
            // for (Entry<String, List<Map<String, String>>> mainData :
            // finalRowDataMap.entrySet()) {

            List< Map< String, String > > inerMapList = finalRowDataMap.get( rowOrderKey );

            rowInt++;
            Row row = sheet.createRow( rowInt );
            sheet.addMergedRegion( new CellRangeAddress( rowInt, rowInt, 0, headers.size() ) );
            Cell headingCol = row.createCell( 0 );
            headingCol.setCellValue( rowOrderKey );
            headingCol.setCellStyle( titalStyle );

            for ( Map< String, String > innerDataColMap : inerMapList ) {

                rowInt++;
                Row rowActualDataValue = sheet.createRow( rowInt );
                int cellInt = 0;
                for ( String fetchFromHeader : headers ) {

                    Cell dataCell = rowActualDataValue.createCell( cellInt );
                    dataCell.setCellValue( innerDataColMap.get( fetchFromHeader ) );
                    dataCell.setCellStyle( rowStyle );
                    cellInt++;

                    // wirth set
                    sheet.autoSizeColumn( cellInt );
                    sheet.setColumnWidth( cellInt, sheet.getColumnWidth( cellInt ) );
                }
            }
        }

        FileOutputStream outputStream = new FileOutputStream( outFile );
        workbook.write( outputStream );
        workbook.close();

        return outFile;
    }

}
