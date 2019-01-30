package com.bookmarks;

import com.aspose.words.*;
import com.aspose.words.Shape;
import com.help.TestFile;
import org.testng.annotations.Test;

import java.awt.*;

/**
 * @Auther: zch
 * @Date: 2019/1/30 17:11
 * @Description:
 */
public class WorkingWithShapes {
    private  String dataDir =  TestFile.getTestDataParentDir(this.getClass());

    /**
     * 插入图片
     * @throws Exception
     */
    @Test
    public void insertShapeUsingDocumentBuilder()throws Exception{
        // ExStart:InsertShapeUsingDocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //Free-floating shape insertion. 自由浮动形状插入。
        Shape shape = builder.insertImage(dataDir + "Images\\dotnet-logo.png",
                RelativeHorizontalPosition.PAGE, 100,
                RelativeVerticalPosition.PAGE, 0,
                50, 50,
                WrapType.NONE);

        shape.setRotation(30.0);

        builder.writeln();

        //Inline shape insertion. 内嵌形状插入。
        shape =  builder.insertImage(dataDir + "Images\\Test_636_852.gif",
                RelativeHorizontalPosition.PAGE, 100,
                RelativeVerticalPosition.PAGE, 100,
                50, 50,
                WrapType.NONE);
        shape.setRotation(30.0);

        OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.DOCX);
        // "Strict" or "Transitional" compliance allows to save shape as DML.
        so.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);

        dataDir = dataDir + "Shape_InsertShapeUsingDocumentBuilder_out.docx";

        // Save the document to disk.
        doc.save(dataDir, so);
        // ExEnd:InsertShapeUsingDocumentBuilder
        System.out.println("\nInsert Shape successfully using DocumentBuilder.\nFile saved at " + dataDir);
    }

    /**
     * 插入字体图片
     * @throws Exception
     */
    @Test
    public void setShapeLayoutInCell()throws Exception{
        Document doc = new Document(dataDir + "MyBookmark.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
        watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
//        watermark.isLayoutInCell(false); // Display the shape outside of table cell if it will be placed into a cell.

        watermark.setWidth(300);
        watermark.setHeight(70);
        watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
        watermark.setVerticalAlignment(VerticalAlignment.CENTER);

        watermark.setRotation(-40);
        watermark.getFill().setColor(Color.GRAY);
        watermark.setStrokeColor(Color.GRAY);

        watermark.getTextPath().setText("watermarkText");
        watermark.getTextPath().setFontFamily("Arial");

        watermark.setName("WaterMark_0");
        watermark.setWrapType(WrapType.NONE);

        Run run = (Run) doc.getChildNodes(NodeType.RUN, true).get(doc.getChildNodes(NodeType.RUN, true).getCount() - 1);

        builder.moveTo(run);
        builder.insertNode(watermark);
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);

        // Save the document to disk.
        dataDir = dataDir + "Shape_IsLayoutInCell_out.docx";
        doc.save(dataDir);
        // ExEnd:SetShapeLayoutInCell
        System.out.println("\nShape's IsLayoutInCell property is set successfully.\nFile saved at " + dataDir);
    }

    /**
     * 添加版式
     * @throws Exception
     */
    @Test
    public void addCornersSnipped()throws Exception{
        // ExStart:AddCornersSnipped
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertImage(dataDir + "Images\\dotnet-logo.png", 50, 50);
        shape.setWidth(100);
        shape.setHeight(100);

        OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.DOCX);
        so.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        dataDir = dataDir + "AddCornersSnipped_out.docx";

        //Save the document to disk.
        doc.save(dataDir, so);
        // ExEnd:AddCornersSnipped
        System.out.println("\nCorner Snip shape is created successfully.\nFile saved at " + dataDir);
    }

    /**
     * 插入一个图标 图形
     * @throws Exception
     */
    @Test
    public void testSingleChartDataPointOfAChartSeries()throws Exception{
        //ExStart:WorkWithSingleChartDataPointOfAChartSeries
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 432, 252);

        // Get first series.
        ChartSeries series0 = shape.getChart().getSeries().get(0);

        // Get second series.
        ChartSeries series1 = shape.getChart().getSeries().get(1);

        ChartDataPointCollection dataPointCollection = series0.getDataPoints();

        // Add data point to the first and second point of the first series.
        ChartDataPoint dataPoint00 = dataPointCollection.add(0);
        ChartDataPoint dataPoint01 = dataPointCollection.add(1);

        // Set explosion.
        dataPoint00.setExplosion(50);

        // Set marker symbol and size.
        dataPoint00.getMarker().setSymbol(MarkerSymbol.CIRCLE);
        dataPoint00.getMarker().setSize(15);

        dataPoint01.getMarker().setSymbol(MarkerSymbol.DIAMOND);
        dataPoint01.getMarker().setSize(20);

        // Add data point to the third point of the second series.
        ChartDataPoint dataPoint12 = series1.getDataPoints().add(2);
        dataPoint12.setInvertIfNegative(true);
        dataPoint12.getMarker().setSymbol(MarkerSymbol.STAR);
        dataPoint12.getMarker().setSize(20);

        doc.save(dataDir + "SingleChartDataPointOfAChartSeries_out.docx");
        //ExEnd:WorkWithSingleChartDataPointOfAChartSeries

    }

}
