package StingRaySharp.Parser.Utility;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;


import org.junit.jupiter.engine.descriptor.DynamicDescendantFilter;

/*

public class WriteStripPDFBox  extends PDFTextStripper{
    
    
   public static List<String> lines = new ArrayList<String>();
    
    public WriteStripPDFBox() {
    }
    
    
    
    public static void main( String[] args ) throws IOException {
        PDDocument document = null;
        String fileName = "apache.pdf";
        try {
            
             document = Loader.loadPDF(new File( "C:\\Users\\noman\\Desktop\\sample_info.pdf" ));
          
            PDFTextStripper stripper = new WriteStripPDFBox();
            stripper.setSortByPosition( true );
            stripper.setStartPage( 0 );
            stripper.setEndPage( document.getNumberOfPages() );
  
            Writer dummy = new OutputStreamWriter(new ByteArrayOutputStream());
            stripper.writeText(document, dummy);
             
            // print lines
           
        }
        finally {
            if( document != null ) {
                document.close();
            }
        }
    }
    
    @Override
    protected void writeString(String str, List<TextPosition> textPositions) {
        lines.add(str);
        System.out.println(  str);
        System.out.println(  "****************8");
        // you may process the line here itself, as and when it is obtained
    }
    
    */


