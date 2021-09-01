package StingRaySharp.Parser;

import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

public class FileGUI {
    
    
    public static void main( String[] args ) throws ClassNotFoundException, InstantiationException, IllegalAccessException, UnsupportedLookAndFeelException {
        UIManager.setLookAndFeel( UIManager.getSystemLookAndFeelClassName() );
        XlsxFileGUI gui = new XlsxFileGUI();
        gui.showGUI(); 
    }
    
    

}
