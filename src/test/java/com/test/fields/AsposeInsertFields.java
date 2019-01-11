package com.test.fields;

import com.aspose.words.*;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:00
 * @Description:字段插入
 */
public class AsposeInsertFields {

    /**
     * 字段插入
     * @throws Exception
     */
    @Test
    public void test() throws Exception{
        final String dataPath = Utils.getDataDir(this.getClass());

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert few page breaks (just for testing) 插入几个分页符（仅用于测试）
        for (int i = 0; i < 5; i++)
            builder.insertBreak(BreakType.PAGE_BREAK);

        // Move DocumentBuilder cursor into the primary footer. 将DocumentBuilder光标移到主页脚中
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);

        // We want to insert a field like this: 我们要插入这样的字段
        // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        Field field = builder.insertField("IF ");
        builder.moveTo(field.getSeparator());
        builder.insertField("PAGE");
        builder.write(" <> ");
        builder.insertField("NUMPAGES");
        builder.write(" \"See Next Page\" \"Last Page\" ");

        // Finally update the outer field to recalcaluate the final value. Doing this will automatically update 最后更新外部字段以重新计算最终值。这样做将自动更新
        // the inner fields at the same time.同时显示内部字段
        field.update();

        doc.save(dataPath + "AsposeFields.docx");
        System.out.println("Aspose Fields Inserted...");

    }

}
