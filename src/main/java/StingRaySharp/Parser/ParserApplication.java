
package StingRaySharp.Parser;

import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

//@SpringBootApplication
public class ParserApplication {

    public static void main( String[] args )
            throws ClassNotFoundException, InstantiationException, IllegalAccessException, UnsupportedLookAndFeelException {
        // SpringApplication.run( ParserApplication.class, args );
        UIManager.setLookAndFeel( UIManager.getSystemLookAndFeelClassName() );
        XlsxFileGUI gui = new XlsxFileGUI();
        gui.showGUI();
    }

}
