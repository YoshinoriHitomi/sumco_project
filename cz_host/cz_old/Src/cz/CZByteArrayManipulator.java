package cz;

/************************************************************
 * Byte Array Manipulator
 * @author  (KPK Co.,Ltd.)
 * @version 1.0 (2003/04/01)package cz;


 ***********************************************************/
public class CZByteArrayManipulator {
    private int
        len,    // length of "byte[] bytes"
        h2,l2,  // high,low pointer in 2-byte
        h4,l4,  // high,low pointer in 4-byte
        h8,l8   // high,low pointer in 8-byte
    ;
    private byte[] bytes;
    //------------------------------------------------------
    byte[] getBytes() { return bytes; }

    //------------------------------------------------------
    CZByteArrayManipulator(boolean bigendian) {
        h2 = bigendian ? 0 : 1; l2 = 1  - h2;
        h4 = h2 << 1;           l4 = l2 << 1;
        h8 = h4 << 1;           l8 = l4 << 1;
    }
    CZByteArrayManipulator(int len, boolean bigendian) {
        this(bigendian);
        this.len = len;
        bytes = new byte[len];
    }
    CZByteArrayManipulator(byte[] bytes, boolean bigendian) {
        this(bigendian);
        this.bytes = bytes;
        len = bytes.length;
    }

    //------------------------------------------------------
    // set n-byte
    void set1(byte by, int i) {
        bytes[i] = by;
    }
    void set1(int in, int i) {
        bytes[i] = (byte)in;
    }
    void set2(int in, int i) {
        bytes[i+h2] = (byte)(in >> 8);
        bytes[i+l2] = (byte)in;
    }
    void set4(int in, int i) {
        set2(in >> 16, i+h4);
        set2(in,       i+l4);
    }
    void set8(long lo, int i) {
        set4((int)(lo >> 32), i+h8);
        set4((int)lo,         i+l8);
    }

    //------------------------------------------------------
    // get n-byte (signed, unsigned)
    byte get1(int i) {
        return bytes[i];
    }
    int getS1(int i) {
        return (int)bytes[i];
    }
    int getU1(int i) {
        return getS1(i) & 0xff;
    }

    int getS2(int i) {
        int ret = (bytes[i+h2] << 8) | (bytes[i+l2] & 0xff);
        return ret;
    }
    int getU2(int i) {
        return getS2(i) & 0xffff;
    }

    int getS4(int i) {
        return (getS2(i+h4) << 16) | getU2(i+l4);
    }
    long getU4(int i) {
        return (long)getS4(i) & 0xffffffffL;
    }

    long getS8(int i) {
        return (getS4(i+h8) << 32) | getU4(i+l8);
    }

    //------------------------------------------------------
    // dump bytes to debug
    void dump() {
        int i = 0;
        try {
            for (;; i++) {
                System.out.print(
                    " " + Integer.toString((int)get1(i), 16)
                );
                if ((i & 7) == 7)
                    System.out.println("");
            }
        } catch (IndexOutOfBoundsException e) {
            System.out.println("");
        }
    }
}
