using System.IO;
using Microsoft.WindowsAPICodePack;
using Microsoft.WindowsAPICodePack.Shell;
using System.Drawing;
using System;



namespace PDF_DIFFERENCE {

    // 2011 Copyright (C) jgr=&jgr, via http://www.theswamp.org
    // 2012 (me): Added code to read PNG Thumbnails from DWG (2013 file format)

    // DWG Thumbnail reader
    internal sealed class ThumbnailReader {
        private ThumbnailReader() {
        }
        static internal Bitmap GetBitmap(string fileName) {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                using (BinaryReader br = new BinaryReader(fs)) {
                    fs.Seek(0xd, SeekOrigin.Begin);
                    fs.Seek(0x14 + br.ReadInt32(), SeekOrigin.Begin);
                    byte bytCnt = br.ReadByte();
                    if (bytCnt <= 1) {
                        return null;
                    }
                    int imageHeaderStart = 0;
                    int imageHeaderSize = 0;
                    byte imageCode = 0;
                    for (short i = 1; i <= bytCnt; i++) {
                        imageCode = br.ReadByte();
                        imageHeaderStart = br.ReadInt32();
                        imageHeaderSize = br.ReadInt32();
                        // BMP Preview (2012 file format)
                        if (imageCode == 2) {
                            // BITMAPINFOHEADER (40 bytes)
                            br.ReadBytes(0xe);
                            //biSize, biWidth, biHeight, biPlanes
                            ushort biBitCount = br.ReadUInt16();
                            br.ReadBytes(4);
                            //biCompression
                            uint biSizeImage = br.ReadUInt32();
                            //br.ReadBytes(0x10); //biXPelsPerMeter, biYPelsPerMeter, biClrUsed, biClrImportant
                            //-----------------------------------------------------
                            fs.Seek(imageHeaderStart, SeekOrigin.Begin);
                            byte[] bitmapBuffer = br.ReadBytes(imageHeaderSize);
                            uint colorTableSize = Convert.ToUInt32(Math.Truncate((biBitCount < 9) ? 4 * Math.Pow(2, biBitCount) : 0));
                            using (MemoryStream ms = new MemoryStream()) {
                                using (BinaryWriter bw = new BinaryWriter(ms)) {
                                    bw.Write(Convert.ToUInt16(0x4d42));
                                    bw.Write(54u + colorTableSize + biSizeImage);
                                    bw.Write(new ushort());
                                    bw.Write(new ushort());
                                    bw.Write(54u + colorTableSize);
                                    bw.Write(bitmapBuffer);
                                    return new Bitmap(ms);
                                }
                            }
                            // PNG Preview (2013 file format)
                        } else if (imageCode == 6) {
                            fs.Seek(imageHeaderStart, SeekOrigin.Begin);
                            using (MemoryStream ms = new MemoryStream()) {
                                //  fs.CopyTo(ms, imageHeaderStart)
                                fs.CopyTo(ms);
                                dynamic img = Image.FromStream(ms);
                                return img;
                            }
                        } else if (imageCode == 3) {
                            return null;
                        }
                    }
                }
            }
            return null;
        }
    }

}