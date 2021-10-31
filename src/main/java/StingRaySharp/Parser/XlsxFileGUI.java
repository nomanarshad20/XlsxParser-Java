
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

import StingRaySharp.Parser.Utility.UccPdfToXlsxParserUtility;
import StingRaySharp.Parser.Utility.GoodWinXlsxParserUtility;

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
    private TextField txtXlsxOne;

    private TextField txtPdfTwo;

    /** The filebytes. */
    byte[] filebytes = null;

    /**
     * Show GUI.
     */
    public void showGUI() {

        jFrame = new JFrame();
        jPanel = new JPanel();
        jPanel.setLayout( null );

        // first Row xlsx
        JLabel xlsxJLable = new JLabel( "GOODWIN :: Select Input Xlsx File" );
        xlsxJLable.setBounds( 10, 10, 190, 25 );

        txtXlsxOne = new TextField( 20 );
        txtXlsxOne.setBounds( 200, 10, 190, 26 );

        JButton browesXlsx = new JButton( "Browse" );
        browesXlsx.setBounds( 400, 10, 80, 28 );

        JButton genrateXlsxOne = new JButton( "  Generate Xlsx/Xlsx  " );
        genrateXlsxOne.setBounds( 500, 10, 150, 28 );

        // second row Pdf
        JLabel pdfJLable = new JLabel( "UCC :: Select Input Pdf File" );
        pdfJLable.setBounds( 10, 40, 190, 25 );

        txtPdfTwo = new TextField( 20 );
        txtPdfTwo.setBounds( 200, 40, 190, 26 );

        JButton browesPdf = new JButton( "Browse" );
        browesPdf.setBounds( 400, 40, 80, 28 );

        JButton genrateXlsxTwo = new JButton( "  Generate Pdf/Xlsx  " );
        genrateXlsxTwo.setBounds( 500, 40, 150, 28 );

        JButton btnClose = new JButton( "Close" );
        btnClose.setBounds( 300, 90, 80, 28 );

        /**
         * Add all components in JPanel
         */
        // one
        jPanel.add( xlsxJLable );
        jPanel.add( txtXlsxOne );
        jPanel.add( browesXlsx );
        jPanel.add( genrateXlsxOne );

        // two
        jPanel.add( pdfJLable );
        jPanel.add( txtPdfTwo );
        jPanel.add( browesPdf );
        jPanel.add( genrateXlsxTwo );

        //

        jPanel.add( btnClose );

        jPanel.setSize( 700, 200 );
        jFrame.add( jPanel );

        /**
         * Set JFrame properties
         */
        jFrame.setLocation( 650, 300 );
        jFrame.pack();
        jFrame.setSize( 700, 200 );
        jFrame.setTitle( "Parser-Xlsx-Pdf( Telos Legal Corp )" );
        jFrame.setDefaultCloseOperation( JFrame.EXIT_ON_CLOSE );
        jFrame.setVisible( true );

        browesXlsx.addActionListener( new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent arg0 ) {

                openForLicense();

            }
        } );

        browesPdf.addActionListener( new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent arg0 ) {

                openForPfx();

            }
        } );

        genrateXlsxOne.addActionListener( new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent arg0 ) {
                try {
                    // && ( !txtLicenseTemplate.getText().contains( ".xlsx" ) || !txtLicenseTemplate.getText().contains( ".XLSX" ) )
                    if ( txtXlsxOne.getText() != null && !txtXlsxOne.getText().isEmpty() ) {
                        try {

                            File file = new File( txtXlsxOne.getText() );
                            String outFile = file.getParent() + File.separator + file.getName() + "_" + new Date().getTime() + "_.xlsx";

                            String outputFilePath = GoodWinXlsxParserUtility.readXlsx( txtXlsxOne.getText(), outFile );

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

        genrateXlsxTwo.addActionListener( new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent arg0 ) {
                try {
                    // && ( !txtLicenseTemplate.getText().contains( ".xlsx" ) || !txtLicenseTemplate.getText().contains( ".XLSX" ) )
                    if ( txtPdfTwo.getText() != null && !txtPdfTwo.getText().isEmpty() ) {
                        try {
                            String outputFilePath = UccPdfToXlsxParserUtility.createXlsxFromPdf( txtPdfTwo.getText() );
                            JOptionPane.showMessageDialog( null, FILE_SAVED_MESSAGE + outputFilePath );
                        } catch ( Exception e ) {
                            JOptionPane.showMessageDialog( null, e.getMessage() );
                        }

                    } else {
                        JOptionPane.showMessageDialog( null, "Please select Pdf File" );
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
            txtXlsxOne.setText( file.getAbsolutePath() );
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
            txtPdfTwo.setText( file.getAbsolutePath() );
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