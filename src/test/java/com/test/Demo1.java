package com.test;

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.*;
import java.util.UUID;

/**
 * @Auther: zch
 * @Date: 2019/1/10 16:29
 * @Description:
 */
public class Demo1 {
    String dir = "D:\\temp\\word\\";

    /**
     * 第一个例子
     *
     * @throws Exception
     */
    @Test
    public void testA() throws Exception {
        Document document = new Document();
        document.getMailMerge().execute(new String[]{"fullName"}, new Object[]{"james"});
        document.save(String.format("%s%s%s", dir, UUID.randomUUID().toString().substring(1, 7), ".docx"));
    }

    @Test
    public void testB() throws Exception {
        //新建一个空白的文档
        Document doc = new Document();
        /*这里面的`builder`相当于一个画笔，提前给他规定样式，然后他就能根据你的要求画出你想画的Word。这里的画笔使用的是就近原则，当上面没有定义了builder的时候，
        会使用默认的格式，当上面定义了某个格式的时候，使用最近的一个（即最后一个改变的样式）*/
        DocumentBuilder builder = new DocumentBuilder(doc);
        //设定Word页面的样式
        //A4纸
        builder.getPageSetup().setPaperSize(PaperSize.A4);
        //方向
        builder.getPageSetup().setOrientation(Orientation.PORTRAIT);
        //垂直对准
        builder.getPageSetup().setVerticalAlignment(PageVerticalAlignment.TOP);
        //页面左边距
        builder.getPageSetup().setLeftMargin(42);
        //页面右边距
        builder.getPageSetup().setRightMargin(42);
        /*关于页面的设置，基本都在PageSetup中，根据需要和正常的名字，基本都可以猜出来*/

        //获取ParagraphFormat对象，关于行的样式基本都在这里
        ParagraphFormat ph = builder.getParagraphFormat();
        //文字对齐方式
        ph.setAlignment(ParagraphAlignment.CENTER);
        // 单倍行距 = 12 ， 1.5 倍 = 18
        ph.setLineSpacing(12);

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
        //添加文字
        builder.write("今天天气不错啊!");
        //添加回车
        builder.writeln();
        //添加文字后回车
        builder.writeln("是啊!是个学习的日子!");
        builder.writeln("https://www.cnblogs.com/itljf/p/10179028.html");
        /*注意：`builder`在`Write`的时候，默认会使用上面规定的格式，除非你在使用`Write`前更新画笔的格式，所以，当你在做样式很多的Word的时候注意更改画笔的格式。*/

        //添加图片
        builder.insertImage("http://a.amap.com/jsapi_demos/static/demo-center/icons/poi-marker-default.png");
        //可以控制图片的宽高
        builder.insertImage("http://a.amap.com/jsapi_demos/static/demo-center/icons/poi-marker-default.png", 20, 20);
        //基本是这样使用，当然还有是其他很多种的参数，比如Image或Stream等，在使用的时候可以根据需要使用

        //表格使用
        //开始添加表格
        Table table = builder.startTable();
        //开始添加第一行，并设置表格行高
        RowFormat rowf = builder.getRowFormat();
        rowf.setHeight(40);
        /*....这里rowf可以有很多的设置*/
        //设置单元格是否水平合并，None为不合并
        builder.getCellFormat().setHorizontalMerge(CellMerge.NONE);
        //设置单元格是否垂直合并，None为不合并
        builder.getCellFormat().setVerticalMerge(CellMerge.NONE);
        //设置单元格宽
        builder.getCellFormat().setWidth(40);
        //单元格垂直对齐方向
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        //单元格水平对齐方向
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);
        //单元格内文字设为多行（默认为单行，会影响单元格宽）
        builder.getCellFormat().setFitText(true);
        builder.write("这是第一行第一个单元格");
        builder.insertCell().getCellFormat().setWidth(12);
        builder.getCellFormat().setWidth(2);//当不需要规定这个单元格的宽度的时候，设置成-1，会是自动宽度
        builder.write("这是第一行第二个单元格");
        //结束第一行
        builder.endRow();
        //结束表格
        builder.endTable();
        //设置这个表格的上下左右，内部水平，垂直的线为白色（当背景为白色的时候就相当于隐藏边框了）
        table.setBorder(BorderType.LEFT, LineStyle.DOUBLE, 1, Color.WHITE, false);
        table.setBorder(BorderType.TOP, LineStyle.DOUBLE, 1, Color.WHITE, false);
        table.setBorder(BorderType.RIGHT, LineStyle.DOUBLE, 1, Color.WHITE, false);
        table.setBorder(BorderType.BOTTOM, LineStyle.DOUBLE, 1, Color.WHITE, false);
        table.setBorder(BorderType.VERTICAL, LineStyle.DOUBLE, 1, Color.WHITE, false);
        /*注意：最重要的是不用忘记开始表格，开始一行，结束一行，结束表格
        里面的设置可以根据个人需要修改，也可以不写使用默认的*/

        doc.save(String.format("%s%s%s", dir, UUID.randomUUID().toString().substring(1, 7), ".docx"));
    }
}
