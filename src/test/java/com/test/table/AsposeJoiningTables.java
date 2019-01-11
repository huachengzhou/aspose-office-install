package com.test.table;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:41
 * @Description:
 */
public class AsposeJoiningTables {

    @Test
    public void test() throws Exception {
        final String dataPath = Utils.getDataDir(this.getClass());
        // Load the document.
        Document doc = new Document(dataPath + "tableDoc.doc");

        // Get the first and second table in the document. 获取文档中的第一和第二个表
        // The rows from the second table will be appended to the end of the first table. 第二个表中的行将附加到第一个表的末尾。
        Table firstTable = (Table)doc.getChild(NodeType.TABLE, 0, true);
        Table secondTable = (Table)doc.getChild(NodeType.TABLE, 1, true);

        // Append all rows from the current table to the next. 将当前表中的所有行追加到下一行
        //由于表的设计，即使是具有不同单元格计数和宽度的表也可以合并为一个表。
        // Due to the design of tables even tables with different cell count and widths can be joined into one table.
        while (secondTable.hasChildNodes())
            firstTable.getRows().add(secondTable.getFirstRow());

        // Remove the empty table container.取出空的表容器
        secondTable.remove();

        doc.save(dataPath + "AsposeJoinTables.doc");

        System.out.println("Done.");
    }

}
