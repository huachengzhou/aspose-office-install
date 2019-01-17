package org.test;

import com.aspose.words.*;
import com.help.TestFile;
import org.testng.annotations.Test;

import java.awt.*;
import java.util.UUID;

/**
 * @Author noatn
 * @Description
 * @createDate 2019/1/16
 **/
public class Demo2Table {

    @Test
    public void testE() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();

        //自动计算行数(第一行)
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.endRow();

        //第二行
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.endRow();

        //第三行
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.insertCell();
        //这设置下单元格样式
        builder.getCellFormat().getBorders().getLeft().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getRight().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getTop().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getBottom().setLineWidth(4.0);
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.endRow();

        builder.endTable();

        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), UUID.randomUUID().toString().substring(1, 12), ".docx"));
    }

    /**
     * 合并
     */
    @Test
    public void testMerge() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();

        //自动计算行数(第一行)
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.endRow();

        //第二行
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.endRow();

        //第三行
        builder.insertCell();
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.insertCell();
        //这设置下单元格样式
        builder.getCellFormat().getBorders().getLeft().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getRight().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getTop().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getBottom().setLineWidth(4.0);
        builder.writeln(String.format("%s%s%s", UUID.randomUUID().toString(), ",", UUID.randomUUID().toString()));
        builder.endRow();


        // We want to merge the range of cells found in between these two cells.
        Cell cellStartRange = table.getRows().get(0).getCells().get(0); //第1行第1列
        Cell cellEndRange = table.getRows().get(1).getCells().get(0); //第2行第1列
        // Merge all the cells between the two specified cells into one.
        mergeCells(cellStartRange, cellEndRange, table);
        builder.endTable();

        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), UUID.randomUUID().toString().substring(1, 12), ".docx"));
    }

    private void mergeCells(Cell startCell, Cell endCell, Table parentTable) {
        // Find the row and cell indices for the start and end cell.
        Point startCellPos = new Point(startCell.getParentRow().indexOf(startCell), parentTable.indexOf(startCell.getParentRow()));
        Point endCellPos = new Point(endCell.getParentRow().indexOf(endCell), parentTable.indexOf(endCell.getParentRow()));
        // Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell.
        Rectangle mergeRange = new Rectangle(
                Math.min(startCellPos.x, endCellPos.x),
                Math.min(startCellPos.y, endCellPos.y),
                Math.abs(endCellPos.x - startCellPos.x) + 1,
                Math.abs(endCellPos.y - startCellPos.y) + 1
        );

        for (Row row : parentTable.getRows()) {
            for (Cell cell : row.getCells()) {
                Point currentPos = new Point(row.indexOf(cell), parentTable.indexOf(row));

                // Check if the current cell is inside our merge range then merge it.
                if (mergeRange.contains(currentPos)) {
                    if (currentPos.x == mergeRange.x)
                        cell.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
                    else
                        cell.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);

                    if (currentPos.y == mergeRange.y)
                        cell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
                    else
                        cell.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);
                }
            }
        }
    }

}
