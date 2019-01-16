package org.test;

import com.aspose.words.*;
import com.aspose.words.Shape;
import com.help.TestFile;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.jsoup.Jsoup;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.io.File;
import java.util.UUID;

/**
 * @Description
 * @createDate 2019/1/15
 **/
public class Demo1 {

    @Test
    public void testA()throws Exception{

        final String dataPath = TestFile.getTestDataParentDir(this.getClass());

        HttpClient httpClient = HttpClients.createDefault();

        HttpPost httpPost = new HttpPost("https://www.cnblogs.com/");
        httpPost.setHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.67 Safari/537.36");
        HttpResponse httpResponse = null;

        httpResponse = httpClient.execute(httpPost);
        if (httpResponse == null) {
            return;
        }
        HttpEntity httpEntity = httpResponse.getEntity();
        if (httpEntity == null) {
            return;
        }
        String webHtml = EntityUtils.toString(httpEntity, "utf-8");
        if (StringUtils.isEmpty(webHtml)) {
            return;
        }
        org.jsoup.nodes.Document document = Jsoup.parse(webHtml);
        org.jsoup.select.Elements post_list = document.select("#post_list .post_item .post_item_body > p");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        //获取Font对象，关于文字的大小，颜色，字体等等基本都在这个里面
        com.aspose.words.Font font = builder.getFont();
        //字体大小
        font.setSize(22);
        //是否粗体
        font.setBold(false);
        //下划线样式，None为无下划线
        font.setUnderline(Underline.DASH);
        //字体颜色
        font.setColor(Color.black);
        //设置字体
        font.setNameFarEast("宋体");
        for (org.jsoup.nodes.Element ele:post_list){
            builder.writeln(ele.text());
        }

        //插入条形码
        int numPages = 4;
        insertBarcodeIntoFooter(builder, doc.getFirstSection(), 1, HeaderFooterType.FOOTER_PRIMARY,dataPath);
        int i = 1;
        while (i < numPages)
        {
            Section cloneSection = (Section) doc.getFirstSection().deepClone(false);
            cloneSection.getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
            doc.appendChild(cloneSection);
            insertBarcodeIntoFooter(builder, cloneSection, i, HeaderFooterType.FOOTER_PRIMARY,dataPath);
            i += 1;
        }

        //插入水印
        insertWatermarkText(doc, "CONFIDENTIAL");

        //插入一个文档
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);

        builder2.writeln("doc2");
        //书签设置
        String bookmark = UUID.randomUUID().toString();
        builder2.startBookmark(bookmark);
        builder2.writeln("这是一个书签");
        builder2.endBookmark(bookmark);
        doc.appendDocument(doc2,ImportFormatMode.KEEP_SOURCE_FORMATTING);
        PageSetup pageSetup = builder2.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setBottomMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setLeftMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setRightMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setHeaderDistance(ConvertUtil.inchToPoint(0.2));
        pageSetup.setFooterDistance(ConvertUtil.inchToPoint(0.2));


        doc.save(String.format("%s%s%s",dataPath, UUID.randomUUID().toString().substring(1,9),".docx"));
    }

    private void insertBarcodeIntoFooter(DocumentBuilder builder, Section section,
                                         int pageId, int footerType,String dataPath) throws Exception {
        // Move to the footer type in the specific section.
        builder.moveToSection(section.getDocument().indexOf(section));
        builder.moveToHeaderFooter(footerType);

        // Insert the barcode, then move to the next line and insert the ID
        // along with the page number.
        // Use pageId if you need to insert a different barcode on each page. 0
        // = First page, 1 = Second page etc.
        builder.insertImage(ImageIO.read(new File(dataPath + "barcode.png")));
        builder.writeln();
        builder.write("1234567890");
        builder.insertField("PAGE");

        // Create a right aligned tab at the right margin.
        double tabPos = section.getPageSetup().getPageWidth()
                - section.getPageSetup().getRightMargin() - section.getPageSetup().getLeftMargin();
        builder.getCurrentParagraph().getParagraphFormat().getTabStops()
                .add(new TabStop(tabPos, TabAlignment.RIGHT, TabLeader.NONE));

        // Move to the right hand side of the page and insert the page and page
        // total.
        builder.write(ControlChar.TAB);
        builder.insertField("PAGE");
        builder.write(" of ");
        builder.insertField("NUMPAGES");
    }

    /**
     * Inserts a watermark into a document.(在文档中插入水印)
     *
     * @param doc           The input document.
     * @param watermarkText Text of the watermark.
     */
    private void insertWatermarkText(Document doc, String watermarkText) throws Exception {
        // Create a watermark shape. This will be a WordArt shape.
        // You are free to try other shape types as watermarks.
        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);

        // Set up the text of the watermark.
        watermark.getTextPath().setText(watermarkText);
        watermark.getTextPath().setFontFamily("Arial");
        watermark.setWidth(500);
        watermark.setHeight(100);
        // Text will be directed from the bottom-left to the top-right corner.
        watermark.setRotation(-40);
        // Remove the following two lines if you need a solid black text.
        watermark.getFill().setColor(Color.GRAY); // Try LightGray to get more Word-style watermark
        watermark.setStrokeColor(Color.GRAY); // Try LightGray to get more Word-style watermark

        // Place the watermark in the page center.
        watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        watermark.setWrapType(WrapType.NONE);
        watermark.setVerticalAlignment(VerticalAlignment.CENTER);
        watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);

        // Create a new paragraph and append the watermark to this paragraph.
        Paragraph watermarkPara = new Paragraph(doc);
        watermarkPara.appendChild(watermark);

        // Insert the watermark into all headers of each document section.
        for (Section sect : doc.getSections()) {
            // There could be up to three different headers in each section, since we want
            // the watermark to appear on all pages, insert into all headers.
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY);
            // insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_FIRST);
            // insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_EVEN);
        }
    }

    private void insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, int headerType) throws Exception {
        HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);

        if (header == null) {
            // There is no header of the specified type in the current section, create it.
            header = new HeaderFooter(sect.getDocument(), headerType);
            sect.getHeadersFooters().add(header);
        }

        // Insert a clone of the watermark into the header.
        header.appendChild(watermarkPara.deepClone(true));
    }

}
