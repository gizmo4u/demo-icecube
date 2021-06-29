package com.chloesoft;

import java.awt.image.BufferedImage;
import java.io.File;

import javax.imageio.ImageIO;

import com.spire.doc.Document;
import com.spire.doc.documents.ImageType;

/**
 * Hello world!
 */
public final class App {
    private App() {
    }

    /**
     * Says hello to the world.
     * 
     * @param args The arguments of the program.
     */
    public static void main(String[] args) {
        System.out.println("Hello World!");

        try {
            // create a Document object
            Document doc = new Document();

            // load a Word file
            doc.loadFromFile("/Users/soonystory/Downloads/convert_sample.docx");

            // loop through the pages
            for (int i = 0; i < doc.getPageCount(); i++) {

                // save the specific page to a BufferedImage
                BufferedImage image = doc.saveToImages(i, ImageType.Bitmap);

                // write the image data to a .png file
                File file = new File("/Users/soonystory/Downloads/output/" + String.format(("Img-%d.png"), i));
                ImageIO.write(image, "PNG", file);
            }

            System.out.println("image Done");

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
