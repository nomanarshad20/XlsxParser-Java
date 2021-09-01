
package StingRaySharp.Parser;

import java.awt.TextField;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;

import StingRaySharp.Parser.Utility.XlsxParserUtility;

public class XlsxFileGUI extends JFrame {

    /** The Constant serialVersionUID. */
    private static final long serialVersionUID = 453243424237758907L;

    private static final String EXCEPTION_POPUP_MESSAGE_PRE_FIX = "Exception: ";

    /** The Constant FILE_SAVED_MESSAGE. */
    private static final String FILE_SAVED_MESSAGE = "File Saved at ";

    /** The j frame. */
    JFrame jFrame;

    /** The j panel. */
    JPanel jPanel;

    /** The txt private key file. */
    private TextField txtLicenseTemplate;

    private TextField txtPrivateKeyFile;

    /** The filebytes. */
    byte[] filebytes = null;

    /**
     * Show GUI.
     */
    public void showGUI() {

        jFrame = new JFrame();
        jPanel = new JPanel();
        jPanel.setLayout( null );

        /**
         * First Row
         */
        JLabel lblLicenseTemplate = new JLabel( "Select Input Xlsx File" );
        lblLicenseTemplate.setBounds( 10, 10, 190, 25 );

        txtLicenseTemplate = new TextField( 20 );
        txtLicenseTemplate.setBounds( 200, 10, 190, 26 );

        JButton btnBrowseLicenseTemplate = new JButton( "Browse" );
        btnBrowseLicenseTemplate.setBounds( 400, 10, 80, 28 );

        /**
         * Second Row
         */

        /*    
        JLabel lblPrivateKeyFile = new JLabel( "Select Output Xlsx File Path" );
        lblPrivateKeyFile.setBounds( 10, 40, 190, 25 );
        
        txtPrivateKeyFile = new TextField( 20 );
        txtPrivateKeyFile.setBounds( 200, 40, 190, 26 );
        
        JButton btnBrowsePrivateKeyFile = new JButton( "Browse" );
        btnBrowsePrivateKeyFile.setBounds( 400, 40, 80, 28 );
        
        */

        /**
         * 3rd Row
         */
        JButton btnSignLicense = new JButton( "  Generate Xlsx  " );
        btnSignLicense.setBounds( 130, 80, 120, 28 );

        JButton btnClose = new JButton( "Close" );
        btnClose.setBounds( 300, 80, 80, 28 );

        /**
         * Add all components in JPanel
         */
        jPanel.add( lblLicenseTemplate );
        jPanel.add( txtLicenseTemplate );
        jPanel.add( btnBrowseLicenseTemplate );
        // jPanel.add( lblPrivateKeyFile );
        // jPanel.add( txtPrivateKeyFile );
        // jPanel.add( btnBrowsePrivateKeyFile );
        jPanel.add( btnSignLicense );
        jPanel.add( btnClose );

        jPanel.setSize( 500, 250 );
        jFrame.add( jPanel );

        /**
         * Set JFrame properties
         */
        jFrame.setLocation( 750, 300 );
        jFrame.pack();
        jFrame.setSize( 500, 250 );
        jFrame.setTitle( "Xlsx Generator ( Telos Legal Corp )" );
        jFrame.setDefaultCloseOperation( JFrame.EXIT_ON_CLOSE );
        jFrame.setVisible( true );

        btnBrowseLicenseTemplate.addActionListener( new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent arg0 ) {

                openForLicense();

            }
        } );

        /*      
        btnBrowsePrivateKeyFile.addActionListener( new ActionListener() {
        
            @Override
            public void actionPerformed( ActionEvent arg0 ) {
        
                openForPfx();
        
            }
        } );
        */

        btnSignLicense.addActionListener( new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent arg0 ) {
                try {
                    // && ( !txtLicenseTemplate.getText().contains( ".xlsx" ) || !txtLicenseTemplate.getText().contains( ".XLSX" ) )
                    if ( txtLicenseTemplate.getText() != null && !txtLicenseTemplate.getText().isEmpty() ) {
                        try {

                            File file = new File( txtLicenseTemplate.getText() );
                            String outFile = file.getParent() + File.separator + file.getName() + "_" + new Date().getTime() + "_.xlsx";

                            String outputFilePath = XlsxParserUtility.readXlsx( txtLicenseTemplate.getText(), outFile );

                            JOptionPane.showMessageDialog( null, FILE_SAVED_MESSAGE + outputFilePath );
                        } catch ( Exception e ) {
                            JOptionPane.showMessageDialog( null, e.getMessage() );
                        }

                    } else {
                        JOptionPane.showMessageDialog( null, "Please select Xlsx File" );
                    }

                } catch ( Exception e ) {
                    JOptionPane.showMessageDialog( null, EXCEPTION_POPUP_MESSAGE_PRE_FIX + e.getMessage() );
                }
            }
        } );
        btnClose.addActionListener( new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent arg0 ) {
                System.exit( 0 );// NOSONAR
            }
        } );
    }

    /**
     * Open for license.
     */
    private void openForLicense() {
        JFileChooser chooser = new JFileChooser();
        int ret = chooser.showOpenDialog( null );
        if ( ret != JFileChooser.CANCEL_OPTION ) {
            File file = chooser.getSelectedFile();
            txtLicenseTemplate.setText( file.getAbsolutePath() );
        }
    }

    /**
     * Open for pfx.
     */
    private void openForPfx() {
        JFileChooser chooser = new JFileChooser();
        int ret = chooser.showOpenDialog( null );
        if ( ret != JFileChooser.CANCEL_OPTION ) {
            File file = chooser.getSelectedFile();
            txtPrivateKeyFile.setText( file.getAbsolutePath() );
        }
    }

    /**
     * Saves the bytes to the chosen path.
     *
     * @param signedXmlDocument
     *            the a signed xml document
     * @throws IOException
     */
    private void save( String path ) throws IOException {
        System.out.println( path );
        JOptionPane.showMessageDialog( null, FILE_SAVED_MESSAGE + path );
    }

}