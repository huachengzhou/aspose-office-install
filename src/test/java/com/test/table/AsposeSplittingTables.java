package com.test.table;

import com.aspose.words.*;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:47
 * @Description:拆分表
 */
public class AsposeSplittingTables {

    @Test
    public void test() throws Exception {
        final String dataPath = Utils.getDataDir(this.getClass());
        // Load the document.
        Document doc = new Document(dataPath + "tableDoc.doc");

        // Get the first table in the document.获取文档中的第一个表。
        Table firstTable = (Table)doc.getChild(NodeType.TABLE, 0, true);

        // We will split the table at the third row (inclusive). 我们将在第三行（包括第三行）拆分表。
        Row row = firstTable.getRows().get(2);

        // Create a new container for the split table.为拆分表创建新容器。
        Table table = (Table)firstTable.deepClone(false);

        // Insert the container after the original.将容器插入原始容器之后。
        firstTable.getParentNode().insertAfter(table, firstTable);

        // Add a buffer paragraph to ensure the tables stay apart.添加缓冲区段落以确保表保持分开。
        firstTable.getParentNode().insertAfter(new Paragraph(doc), firstTable);

        Row currentRow;

        do
        {
            currentRow = firstTable.getLastRow();
            table.prependChild(currentRow);
        }
        while (currentRow != row);

        doc.save(dataPath + "AsposeSplitTable.doc");

        System.out.println("Done.");
    }

}
