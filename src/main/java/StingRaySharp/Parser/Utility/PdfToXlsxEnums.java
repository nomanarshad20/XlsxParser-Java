package StingRaySharp.Parser.Utility;
import StingRaySharp.Parser.StateUsAbriviationEnums;

public enum PdfToXlsxEnums {

    TYPE_OF_SEARCH( "Type of Search", "UCC Search Report" ),
    JURIST_FILING_OFFICE( "Jurisdiction/Filing Office", "UCC Search Report" ),
    INDEXED_THROUGH( "Indexed Through", "UCC Search Report" ),
    SUBJECT_SEARCH_NAME( "Subject Search Name", "UCC Search Report" ),
    DOCUMENT_NO( "Document No.", "UCC" ),
    FILED( "Filed", "UCC" ),
    DEBTOR( "Debtor", "UCC" ),
    SECURED_PARTY( "Secured Party", "UCC" ),
    AMENDMENT_TYPE( "Amendment Type", "UCC" ),
    FILE_NO( "File No.", "UCC" ),

    UCC( "UCC", "UCC" );

    PdfToXlsxEnums( String key, String value ) {
        this.key = key;
        this.value = value;
    }

    /** The key. */
    private String key;

    /** The value. */
    private String value;

    /**
     * Gets the key by value.
     *
     * @param matrexValue
     *            the matrex value
     * @return the key by value
     */
    public static String getKeyByValue( String matrexValue ) {
        String key = "";
        for ( PdfToXlsxEnums StateUSEnums : values() ) {
            if ( StateUSEnums.getValue().contains( matrexValue ) ) {
                key = StateUSEnums.getKey();
                break;
            }
        }
        return key;
    }

    /**
     * Gets the value by key.
     *
     * @param matrixKey
     *            the matrix key
     * @return the value by key
     */
    public static String getValueByKey( String matrixKey ) {
        String value = "";
        for ( PdfToXlsxEnums StateUSEnums : values() ) {
            if ( StateUSEnums.getKey().contains( matrixKey ) ) {
                value = StateUSEnums.getValue();
                break;
            }
        }
        return value;
    }

    /**
     * Gets the key.
     *
     * @return the key
     */
    public String getKey() {
        return key;
    }

    /**
     * Sets the key.
     *
     * @param key
     *            the new key
     */
    public void setKey( String key ) {
        this.key = key;
    }

    /**
     * Gets the value.
     *
     * @return the value
     */
    public String getValue() {
        return value;
    }

    /**
     * Sets the value.
     *
     * @param value
     *            the new value
     */
    public void setValue( String value ) {
        this.value = value;
    }

}
