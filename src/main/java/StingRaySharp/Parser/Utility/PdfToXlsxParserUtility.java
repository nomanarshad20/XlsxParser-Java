
package StingRaySharp.Parser.Utility;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfTextExtractor;

public class PdfToXlsxParserUtility {

    public static void main( String args[] ) {

        createXlsxFromPdf( "c:\\users\\noman\\desktop\\sample_info.pdf" );
    }

    public static String createXlsxFromPdf( String inpFile ) {
        File inp = new File( inpFile );
        String outFile = inp.getParent() + File.separator +inp.getName()+"_"+ new Date().getSeconds() + ".xlsx";
        try {
            // Create PdfReader instance.
            PdfReader pdfReader = new PdfReader( inpFile );

            // Get the number of pages in pdf.
            int pages = pdfReader.getNumberOfPages();

            StringBuffer pageContent = new StringBuffer();

            // Iterate the pdf through pages.
            for ( int i = 2; i <= pages; i++ ) {
                pageContent.append( PdfTextExtractor.getTextFromPage( pdfReader, i ) );
                pageContent.append( "\n" );
            }

            // Print the page content on console.
            // System.out.println(pageContent);
            // Close the PdfReader.
            pdfReader.close();

            String[] splitArray = pageContent.toString().split( "\n" );

            String typeOfSearch = null;
            String jurisFillingOffice = null;
            String indexedThrough = null;
            String subjectSearchName = null;

            List< Map< String, String > > recList = new ArrayList<>();

            List< String > linesList = new ArrayList<>();

            Map< Integer, Map< String, String > > mapFinal = new HashMap<>();
            Map< String, String > map = new HashMap<>();

            int index = 0;
            String lastkey = null;
            int UccInt = 0;
            System.out.println( splitArray.length );
            while ( true ) {

                String line = splitArray[ index ];

                System.out.println( line );
                System.out.println( "_______________-" );

                if ( line.contains( PdfToXlsxEnums.DOCUMENT_NO.getKey() ) || line.contains( PdfToXlsxEnums.FILED.getKey() )
                        || line.contains( PdfToXlsxEnums.DEBTOR.getKey() ) || line.contains( PdfToXlsxEnums.SECURED_PARTY.getKey() )
                        || line.contains( PdfToXlsxEnums.AMENDMENT_TYPE.getKey() ) || line.contains( PdfToXlsxEnums.FILE_NO.getKey() )
                        || line.contains( PdfToXlsxEnums.TYPE_OF_SEARCH.getKey() )
                        || line.contains( PdfToXlsxEnums.JURIST_FILING_OFFICE.getKey() )
                        || line.contains( PdfToXlsxEnums.INDEXED_THROUGH.getKey() )
                        || line.contains( PdfToXlsxEnums.SUBJECT_SEARCH_NAME.getKey() ) )

                {

                    if ( line.contains( PdfToXlsxEnums.TYPE_OF_SEARCH.getKey() ) ) {
                        System.out.println( "$$$$$$$$$$$" );
                        System.out.println( index );

                        typeOfSearch = splitArray[ index - 1 ];
                        System.out.println( index );
                        System.out.println( "**************" );
                        // map.put( PdfToXlsxEnums.TYPE_OF_SEARCH.getKey(), splitArray[ index - 1 ] );
                    } else if ( line.contains( PdfToXlsxEnums.JURIST_FILING_OFFICE.getKey() ) ) {
                        jurisFillingOffice = splitArray[ index - 1 ];
                        // map.put( PdfToXlsxEnums.JURIST_FILING_OFFICE.getKey(), splitArray[ index - 1 ] );
                    } else if ( line.contains( PdfToXlsxEnums.INDEXED_THROUGH.getKey() ) ) {
                        line = line.replace( PdfToXlsxEnums.INDEXED_THROUGH.getKey(), "" );
                        indexedThrough = line.trim();
                        // map.put( PdfToXlsxEnums.INDEXED_THROUGH.getKey(), line.trim() );
                    } else if ( line.contains( PdfToXlsxEnums.SUBJECT_SEARCH_NAME.getKey() ) ) {
                        subjectSearchName = splitArray[ index - 1 ];
                        // map.put( PdfToXlsxEnums.SUBJECT_SEARCH_NAME.getKey(), splitArray[ index - 1 ] );
                    } else if ( line.contains( PdfToXlsxEnums.DOCUMENT_NO.getKey() ) ) {
                        if ( map.containsKey( PdfToXlsxEnums.DOCUMENT_NO.getKey() ) ) {
                        } else {
                            line = line.replace( PdfToXlsxEnums.DOCUMENT_NO.getKey(), "" );
                            map.put( PdfToXlsxEnums.DOCUMENT_NO.getKey(), line.trim() );
                        }
                    } else if ( lastkey == null && line.contains( PdfToXlsxEnums.FILED.getKey() ) ) {
                        line = line.replace( PdfToXlsxEnums.FILED.getKey(), "" );
                        map.put( PdfToXlsxEnums.FILED.getKey(), line.trim() );
                    } else if ( line.contains( PdfToXlsxEnums.DEBTOR.getKey() ) ) {
                        line = line.replace( PdfToXlsxEnums.DEBTOR.getKey(), "" );
                        map.put( PdfToXlsxEnums.DEBTOR.getKey(), line.trim() );
                    } else if ( line.contains( PdfToXlsxEnums.SECURED_PARTY.getKey() ) ) {
                        line = line.replace( PdfToXlsxEnums.SECURED_PARTY.getKey(), "" );
                        map.put( PdfToXlsxEnums.SECURED_PARTY.getKey(), line.trim() );
                    } else if ( line.contains( PdfToXlsxEnums.AMENDMENT_TYPE.getKey() ) ) {
                        lastkey = PdfToXlsxEnums.AMENDMENT_TYPE.getKey();
                        line = line.replace( PdfToXlsxEnums.AMENDMENT_TYPE.getKey(), "" );
                        map.put( PdfToXlsxEnums.AMENDMENT_TYPE.getKey(), line.trim() );
                    } else if ( line.contains( PdfToXlsxEnums.FILE_NO.getKey() ) ) {
                        line = line.replace( PdfToXlsxEnums.FILE_NO.getKey(), "" );
                        map.put( PdfToXlsxEnums.FILE_NO.getKey(), line.trim() );
                    } else if ( lastkey.contains( PdfToXlsxEnums.AMENDMENT_TYPE.getKey() )
                            && line.contains( PdfToXlsxEnums.FILED.getKey() ) ) {
                        line = line.replace( PdfToXlsxEnums.FILE_NO.getKey(), "" );
                        map.put( PdfToXlsxEnums.FILED.getKey() + "2", line.trim() );
                    }

                } else if ( line.contains( " UCC" ) || line.contains( "Notice of Federal Tax Lien" ) ) {

                    // add map to list : all col added
                    map.put( PdfToXlsxEnums.TYPE_OF_SEARCH.getKey(), typeOfSearch );
                    map.put( PdfToXlsxEnums.JURIST_FILING_OFFICE.getKey(), jurisFillingOffice );
                    map.put( PdfToXlsxEnums.INDEXED_THROUGH.getKey(), indexedThrough );
                    map.put( PdfToXlsxEnums.SUBJECT_SEARCH_NAME.getKey(), subjectSearchName );
                    recList.add( map );
                    mapFinal.put( UccInt, map );
                    // clear map for next iteration of all map col
                    map = new HashMap<>();

                    lastkey = null;
                    // increment for ucc row index count
                    UccInt++;

                } else if ( line.contains( "We assume no liability with respect to the identity" ) ) {
                    break;
                }

                index++;
            }

            // Map< Integer, Map< String, String > > mapFinal = new HashMap<>();
            ObjectMapper objectMapper = new ObjectMapper();
            String carAsString = objectMapper.writeValueAsString( mapFinal );
            System.out.println( carAsString );

            createXlsxFile( mapFinal, outFile );

        } catch ( Exception e ) {
            e.printStackTrace();
        }
        return outFile;

    }

    private static String createXlsxFile( Map< Integer, Map< String, String > > mapFinal, String outFile ) throws IOException {

        // String OutFileResult = "C:\\Users\\noman\\Desktop\\FIg\\" + new
        // Date().getSeconds() + "_LienSummaryResul" + ".xlsx";

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet( "TelosLegalCorp" );

        List< String > headers = new ArrayList<>();
        // headers.add( "#" );
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
        for ( Entry< Integer, Map< String, String > > outerLayerData : mapFinal.entrySet() ) {

            if ( outerLayerData.getKey() == 0 ) {
                continue;
            }

            rowInt++;
            Row rowActualDataValue = sheet.createRow( rowInt );
            int cellInt = 0;

            // cell index zero
            createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue, cellInt++,
                    outerLayerData.getValue().get( PdfToXlsxEnums.DEBTOR.getKey() ) );

            // cell index
            String juristrinction = outerLayerData.getValue().get( PdfToXlsxEnums.JURIST_FILING_OFFICE.getKey() ) + "\n"
                    + outerLayerData.getValue().get( PdfToXlsxEnums.TYPE_OF_SEARCH.getKey() );
            createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue, cellInt++, juristrinction );

            // cell index
            String docnumAndFiled = outerLayerData.getValue().get( PdfToXlsxEnums.DOCUMENT_NO.getKey() ).split( " " )[ 0 ] + " "
                    + outerLayerData.getValue().get( PdfToXlsxEnums.FILED.getKey() );
            createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue, cellInt++, docnumAndFiled );

            // cell index
            createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue, cellInt++,
                    outerLayerData.getValue().get( PdfToXlsxEnums.SECURED_PARTY.getKey() ) );

            // cell index
            createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue, cellInt++, "" );

            createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue, cellInt++,
                    outerLayerData.getValue().get( PdfToXlsxEnums.INDEXED_THROUGH.getKey() ) );

            // cell index
            createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue, cellInt++, "" );
            // row one prepared

            if ( outerLayerData.getValue().containsKey( PdfToXlsxEnums.AMENDMENT_TYPE.getKey() ) ) {
                // row 2nd started
                rowInt++;

                Row rowActualDataValue2 = sheet.createRow( rowInt );

                // cell index
                String docnumAndFiled2 = outerLayerData.getValue().get( PdfToXlsxEnums.FILE_NO.getKey() ) + " " + outerLayerData.getValue()
                        .get( PdfToXlsxEnums.FILED.getKey() + "2" ).replace( PdfToXlsxEnums.FILED.getKey().trim(), "" );
                createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue2, 2, docnumAndFiled2 );

                // cell index
                createCellAndPopulateValue( sheet, rowStyle, rowActualDataValue2, 4,
                        outerLayerData.getValue().get( PdfToXlsxEnums.AMENDMENT_TYPE.getKey() ) );
            }

        }

        FileOutputStream outputStream = new FileOutputStream( outFile );
        workbook.write( outputStream );
        workbook.close();

        System.out.println( "File saved" );

        return outFile;
    }

    private static void createCellAndPopulateValue( XSSFSheet sheet, XSSFCellStyle rowStyle, Row rowActualDataValue, int cellInt,
            String cellValue ) {
        Cell dataCell = rowActualDataValue.createCell( cellInt );
        dataCell.setCellValue( cellValue );
        dataCell.setCellStyle( rowStyle );
        // With set
        sheet.autoSizeColumn( cellInt );
        sheet.setColumnWidth( cellInt, sheet.getColumnWidth( cellInt ) );

        cellInt++;
    }

}
