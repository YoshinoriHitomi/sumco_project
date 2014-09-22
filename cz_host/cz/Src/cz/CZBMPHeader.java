package cz;

/*
 * BMP File Define
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
*/
public abstract class CZBMPHeader {
    // pointers in BITMAPFILEHEADER
    static final int pbfType          =  0; // (2)
    static final int pbfSize          =  2; // (4)
    static final int pbfReserved1     =  6; // (2)
    static final int pbfReserved2     =  8; // (2)
    static final int pbfOffBits       = 10; // (4)
    static final int bf_size          = 14;
    // pointers in BITMAPINFOHEADER
    static final int pbiSize          =  0; // (4)
    static final int pbiWidth         =  4; // (4)
    static final int pbiHeight        =  8; // (4)
    static final int pbiPlanes        = 12; // (2)
    static final int pbiBitCount      = 14; // (2)
    static final int pbiCompression   = 16; // (4)
    static final int pbiSizeImage     = 20; // (4)
    static final int pbiXPelsPerMeter = 24; // (4)
    static final int pbiYPelsPerMeter = 28; // (4)
    static final int pbiClrUsed       = 32; // (4)
    static final int pbiClrImportant  = 36; // (4)
    static final int bi_size          = 40;

    // members of BITMAPFILEHEADER
    int bfType;          // (2) Type of object; set to "BM"
    int bfSize;          // (4) Size of the file, in bytes.
    int bfReserved1;     // (2) Reserved, set to zero.
    int bfReserved2;     // (2) Reserved, set to zero.
    int bfOffBits;       // (4) Offset, in bytes, from this structure
                         //      to the actual bitmap in the file.

    // members of BITMAPINFOHEADER
    int biSize;          // (4) # of bytes required by this structure.
    int biWidth;         // (4) Width of the image, in pixels.
    int biHeight;        // (4) Height of the image, in pixels.
    int biPlanes;        // (2) # of planes for the target device.
    int biBitCount;      // (2) # of bits per pixel. Must be 1, 4, 8, or 24.
    int biCompression;   // (4) Specifies the compression used.
    int biSizeImage;     // (4) Size of the image, in bytes.
    int biXPelsPerMeter; // (4) Horizontal resolution of the image.
    int biYPelsPerMeter; // (4) Vertical resolution of the image.
    int biClrUsed;       // (4) # of colors in the color table used by the image.
    int biClrImportant;  // (4) # of important colors.

    CZByteArrayManipulator fileheader,infoheader;

    /**
     * String representation.
     */
    public String toString() {
        return
            "bfType          = " + fileToHex2(pbfType)          + "\n" +
            "bfSize          = " + fileToHex4(pbfSize)          + "\n" +
            "bfReserved1     = " + fileToHex2(pbfReserved1)     + "\n" +
            "bfReserved2     = " + fileToHex2(pbfReserved2)     + "\n" +
            "bfOffBits       = " + fileToHex4(pbfOffBits)       + "\n" +
            "biSize          = " + infoToHex4(pbiSize)          + "\n" +
            "biWidth         = " + infoToHex4(pbiWidth)         + "\n" +
            "biHeight        = " + infoToHex4(pbiHeight)        + "\n" +
            "biPlanes        = " + infoToHex2(pbiPlanes)        + "\n" +
            "biBitCount      = " + infoToHex2(pbiBitCount)      + "\n" +
            "biCompression   = " + infoToHex4(pbiCompression)   + "\n" +
            "biSizeImage     = " + infoToHex4(pbiSizeImage)     + "\n" +
            "biXPelsPerMeter = " + infoToHex4(pbiXPelsPerMeter) + "\n" +
            "biYPelsPerMeter = " + infoToHex4(pbiYPelsPerMeter) + "\n" +
            "biClrUsed       = " + infoToHex4(pbiClrUsed)       + "\n" +
            "biClrImportant  = " + infoToHex4(pbiClrImportant);
    }

    //------------------------------------------------------
    String fileToHex2(int i) {
        return toHex2(fileheader.getU2(i));
    }
    //------------------------------------------------------
    String fileToHex4(int i) {
        return toHex4(fileheader.getS4(i));
    }
    //------------------------------------------------------
    String infoToHex2(int i) {
        return toHex2(infoheader.getU2(i));
    }
    //------------------------------------------------------
    String infoToHex4(int i) {
        return toHex4(infoheader.getS4(i));
    }
    //------------------------------------------------------
    String toHex2(int in) {
        String ret = Integer.toString(in & 0xffff, 16);
        while (ret.length() < 4)
            ret = "0" + ret;
        return "    " + ret + " (" + in + ")";
    }
    //------------------------------------------------------
    String toHex4(int in) {
        long lo = (long)in & 0xffffffffL;
        String ret = Long.toString(lo, 16);
        while (ret.length() < 8)
            ret = "0" + ret;
        return ret + " (" + lo + ")";
    }
}
