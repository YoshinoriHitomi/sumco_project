package cz;

/***********************************************************
 * Bitmap File Reader
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)
 ***********************************************************/
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class CZBMPReader extends CZBMPHeader {
    File file;
    FileInputStream fis;
    private int[] rgbquad,pix;
    private int wid,hei,widthbytes;

    //------------------------------------------------------
    public CZBMPReader(File file) {
        this.file = file;
    }

    //------------------------------------------------------
    public void read() throws IOException {
        fis = new FileInputStream(file);
        readFileHeader();
        readInfoHeader();
        readRGBQuad();
        readBitmap();
        fis.close();
    }

    //------------------------------------------------------
    public int[] getPix()  { return pix; }
    public int getWidth()  { return wid; }
    public int getHeight() { return hei; }

    //------------------------------------------------------
    void readFileHeader() throws IOException {
        fileheader = new CZByteArrayManipulator(bf_size,false);
        if (!readFully(fis, fileheader.getBytes()))
            throw new IOException("insufficient BITMAPFILEHEADER");

        bfType    = fileheader.getU2(pbfType);
        bfSize    = fileheader.getS4(pbfSize);
        bfOffBits = fileheader.getS4(pbfOffBits);

        if (bfType != 0x4d42)
            throw new IOException("not bmp file");
    }

    //------------------------------------------------------
    void readInfoHeader() throws IOException {
        infoheader = new CZByteArrayManipulator(bi_size,false);
        if (!readFully(fis, infoheader.getBytes()))
            throw new IOException("insufficient BITMAPINFOHEADER");

        biSize        = infoheader.getS4(pbiSize);
        biWidth       = infoheader.getS4(pbiWidth);
        biHeight      = infoheader.getS4(pbiHeight);
        biBitCount    = infoheader.getU2(pbiBitCount);
        biCompression = infoheader.getS4(pbiCompression);

        if (biSize != bi_size)
            throw new IOException("OS2 bmp?");
        wid = (int)biWidth;
        hei = (int)biHeight;
        widthbytes = ((wid * biBitCount + 31) & ~31) >>> 3;
        if (
            (biBitCount != 1) &&
            (biBitCount != 4) &&
            (biBitCount != 8) &&
            (biBitCount != 24)
        )
            throw new IOException("invalid biBitCount:" + biBitCount);
        if (biCompression != 0)
            throw new IOException("compression not supported");
    }

    //------------------------------------------------------
    void readRGBQuad() throws IOException {
        int rgbquadbytes = (int)bfOffBits - bf_size - bi_size;

        if (rgbquadbytes > 0) {
            byte[] bytesRgbQuad = new byte[rgbquadbytes];
            if (!readFully(fis, bytesRgbQuad))
                throw new IOException("insufficient color table");

            int rgbquadlen = rgbquadbytes >> 2;
            rgbquad = new int[rgbquadlen];

            for (int i=0; i<rgbquadlen; i++) {
                int i4 = i << 2;
                rgbquad[i] = getRGB(bytesRgbQuad,i4);
            }
        }
    }

    //------------------------------------------------------
    void readBitmap() throws IOException {
        int xy = wid * hei;
        pix = new int[xy];
        byte[] bline = new byte[widthbytes];
        int[]  iline = new int[wid+7]; // with padding

        while (xy > 0) {
            xy -= wid;
            if (!readFully(fis, bline))
                throw new IOException("insufficient bitmap");

            switch (biBitCount) {
                case 1: // 2 colors
                    for (int x=0; x<wid;) {
                        int by = (int)bline[x>>>3];
                        for (int i=0; i<8; i++) {
                            iline[x++] = rgbquad[(by >>> 7) & 1];
                            by <<= 1;
                        }
                    }
                    break;

                case 4: // 16 colors
                    for (int x=0; x<wid;) {
                        int by = (int)bline[x>>>1];
                        iline[x++] = rgbquad[(by >>> 4) & 0x0f];
                        iline[x++] = rgbquad[ by        & 0x0f];
                    }
                    break;

                case 8: // 256 colors
                    for (int x=0; x<wid; x++) {
                        iline[x] = rgbquad[(int)bline[x] & 0xff];
                    }
                    break;

                default: //case 24:
                    int x3 = 0;
                    for (int x=0; x<wid; x++) {
                        iline[x] = getRGB(bline,x3);
                        x3 += 3;
                    }
            }
            for (int x=0; x<wid; x++)
                pix[xy+x] = iline[x];
        }
    }

    //------------------------------------------------------
    int getRGB(byte[] bs, int i) {
        return
                                     0xff000000   |
            ( ((int)bs[i+2] << 16) & 0x00ff0000 ) |
            ( ((int)bs[i+1] <<  8) & 0x0000ff00 ) |
            (  (int)bs[i  ]        & 0x000000ff )
        ;
    }

    //------------------------------------------------------
    boolean readFully(InputStream is, byte[] bts) throws IOException {
        return readFully(is, bts, 0, bts.length);
    }

    //------------------------------------------------------
    boolean readFully(
        InputStream is, byte[] bts, int off, int len
    ) throws IOException {

        while (len > 0) {
            int rd = is.read(bts,off,len);
            if (rd < 0)
                return false;
            len -= rd;
            off += rd;
        }
        return true;
    }
}
