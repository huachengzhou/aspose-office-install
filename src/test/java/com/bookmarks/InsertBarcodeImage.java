package com.bookmarks;

import com.aspose.words.*;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/30 15:09
 * @Description:暂无(只是插入了四页空白)
 */
public class InsertBarcodeImage {

    @Test
    public void test() throws Exception {
        String dataDir = TestFile.getTestDataParentDir(this.getClass());

        Document doc = new Document(dataDir + "Image.SampleImages.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The number of pages the document should have.
        int numPages = 4;
        // The document starts with one section, insert the barcode into this existing section.
        InsertBarcodeIntoFooter(builder, doc.getFirstSection(), 1, HeaderFooterType.FOOTER_PRIMARY);
        InsertBarcodeIntoFooter(builder, doc.getFirstSection(), 1, HeaderFooterType.FOOTER_PRIMARY);

        for (int i = 1; i < numPages; i++) {
            // Clone the first section and add it into the end of the document.
            Section cloneSection = (Section) doc.getFirstSection().deepClone(false);
            // cloneSection.getPageSetup().getSectionStart() = SectionStart.NEW_PAGE;
            doc.appendChild(cloneSection);

            // Insert the barcode and other information into the footer of the section.
            InsertBarcodeIntoFooter(builder, cloneSection, i, HeaderFooterType.FOOTER_PRIMARY);
        }
        dataDir = dataDir + "Document_out_.docx";
        // Save the document as a PDF to disk. You can also save this directly to a stream.
        doc.save(dataDir);
        //ExEnd:InsertBarcodeImage
    }

    //ExStart:InsertBarcodeIntoFooter
    private void InsertBarcodeIntoFooter(DocumentBuilder builder, Section section, int pageId, int footerType) {
        // Move to the footer type in the specific section.
        try {
            builder.moveToSection(section.getDocument().indexOf(section));
            builder.moveToHeaderFooter(footerType);

            String dataDir = TestFile.getTestDataParentDir(this.getClass());

            // Insert the barcode, then move to the next line and insert the ID along with the page number.
            // Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.
            builder.insertImage(dataDir + "Barcode1.png");
            builder.writeln();
            builder.write("1234567890");
            builder.insertField("PAGE");

            // Create a right aligned tab at the right margin.
            double tabPos = section.getPageSetup().getPageWidth() - section.getPageSetup().getRightMargin() - section.getPageSetup().getLeftMargin();
            builder.getCurrentParagraph().getParagraphFormat().getTabStops().add(new TabStop(tabPos, TabAlignment.RIGHT, TabLeader.NONE));

            // Move to the right hand side of the page and insert the page and page total.
            builder.write(ControlChar.TAB);
            builder.insertField("PAGE");
            builder.write(" of ");
            builder.insertField("NUMPAGES");
        } catch (Exception x) {

        }
    }
    //ExEnd:InsertBarcodeIntoFooter

}
