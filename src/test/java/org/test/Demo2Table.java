package org.test;

import com.aspose.words.*;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.help.TestFile;
import com.red.tool.other.MergeCellModel;
import org.apache.commons.collections.CollectionUtils;
import org.testng.annotations.Test;

import java.awt.*;
import java.util.*;
import java.util.List;

/**
 * @Author zch
 * @Description
 * @createDate 2019/1/16
 **/
public class Demo2Table {

    @Test
    public void testB() throws Exception {
        List<String> stringList = new ArrayList<String>(12);
        for (int i = 0; i < 12; i++) {
            stringList.add(UUID.randomUUID().toString());
        }
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("房屋所有权登记状况表");
        for (String s : stringList) {
            Set<MergeCellModel> mergeCellModelList = Sets.newHashSet();
            Table table = builder.startTable();
            for (int j = 0; j < 7; j++) {
                switch (j) {
                    case 0:
                        for (int k = 0; k < 4; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("权益状况");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    builder.writeln("土地权益类型");
                                    break;
                                case 3:
                                    builder.insertCell();
                                    builder.writeln("");
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 1:
                        for (int k = 0; k < 4; k++) {
                            switch (k) {
                                case 1:
                                    builder.insertCell();
                                    builder.writeln("土地管制情况");
                                    break;
                                case 3:
                                    builder.insertCell();
                                    builder.writeln("");
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 2:
                        for (int k = 0; k < 4; k++) {
                            switch (k) {
                                case 1:
                                    builder.insertCell();
                                    builder.writeln("土地他项权力");
                                    break;
                                case 3:
                                    builder.insertCell();
                                    builder.writeln("");
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 3:
                        for (int k = 0; k < 4; k++) {
                            switch (k) {
                                case 1:
                                    builder.insertCell();
                                    builder.writeln("其他特殊情况");
                                    break;
                                case 3:
                                    builder.insertCell();
                                    builder.writeln("");
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 4:
                        for (int k = 0; k < 4; k++) {
                            switch (k) {
                                case 1:
                                    builder.insertCell();
                                    builder.writeln("房屋所有权");
                                    break;
                                case 3:
                                    builder.insertCell();
                                    builder.writeln("");
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 5:
                        for (int k = 0; k < 4; k++) {
                            switch (k) {
                                case 1:
                                    builder.insertCell();
                                    builder.writeln("出租或占用情况");
                                    break;
                                case 3:
                                    builder.insertCell();
                                    builder.writeln("");
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 6:
                        for (int k = 0; k < 4; k++) {
                            switch (k) {
                                case 1:
                                    builder.insertCell();
                                    builder.writeln("物业管理情况");
                                    break;
                                case 3:
                                    builder.insertCell();
                                    builder.writeln("");
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    default:
                        break;
                }
                mergeCellModelList.add(new MergeCellModel(j,1,j,2));
                mergeCellModelList.add(new MergeCellModel(0,0,6,0));
            }
            if (CollectionUtils.isNotEmpty(mergeCellModelList)) {
                for (MergeCellModel mergeCellModel : mergeCellModelList) {
                    Cell cellStartRange = table.getRows().get(mergeCellModel.getStartRowIndex()).getCells().get(mergeCellModel.getStartColumnIndex());
                    Cell cellEndRange = table.getRows().get(mergeCellModel.getEndRowIndex()).getCells().get(mergeCellModel.getEndColumnIndex());
                    mergeCells(cellStartRange, cellEndRange, table);
                }
            }
            builder.endTable();
        }
        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), "testB", ".docx"));
    }


    @Test
    public void testD() throws Exception {
        List<String> stringList = new ArrayList<String>(12);
        for (int i = 0; i < 12; i++) {
            stringList.add(UUID.randomUUID().toString());
        }
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("估价土地实体状况表");
        for (String s : stringList) {
            Set<MergeCellModel> mergeCellModelList = Sets.newHashSet();
            Table table = builder.startTable();
            //行
            for (int j = 0; j < 8; j++) {
                switch (j) {
                    case 0:
                        for (int k = 0; k < 5; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("估价对象名称");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    builder.writeln(s);
                                    break;
                                case 4:
                                    builder.insertCell();
                                    builder.writeln("备注");
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 1:
                        for (int k = 0; k < 5; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("用途及级别");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    break;
                                case 4:
                                    builder.insertCell();
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 2:
                        for (int k = 0; k < 5; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("东至,南至,西至,北至");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    break;
                                case 4:
                                    builder.insertCell();
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 3:
                        for (int k = 0; k < 5; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("面积");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    break;
                                case 4:
                                    builder.insertCell();
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 4:
                        for (int k = 0; k < 5; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("形状");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    break;
                                case 4:
                                    builder.insertCell();
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 5:
                        for (int k = 0; k < 5; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("地形、地势、工程地质");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    break;
                                case 4:
                                    builder.insertCell();
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 6:
                        for (int k = 0; k < 5; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("开发程度");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    break;
                                case 4:
                                    builder.insertCell();
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    case 7:
                        for (int k = 0; k < 5; k++) {
                            switch (k) {
                                case 0:
                                    builder.insertCell();
                                    builder.writeln("其它");
                                    break;
                                case 1:
                                    builder.insertCell();
                                    break;
                                case 4:
                                    builder.insertCell();
                                    break;
                                default:
                                    builder.insertCell();
                                    break;
                            }
                        }
                        builder.endRow();
                        break;
                    default:
                        break;
                }
                mergeCellModelList.add(new MergeCellModel(j, 1, j, 3));
            }
            if (CollectionUtils.isNotEmpty(mergeCellModelList)) {
                for (MergeCellModel mergeCellModel : mergeCellModelList) {
                    Cell cellStartRange = table.getRows().get(mergeCellModel.getStartRowIndex()).getCells().get(mergeCellModel.getStartColumnIndex());
                    Cell cellEndRange = table.getRows().get(mergeCellModel.getEndRowIndex()).getCells().get(mergeCellModel.getEndColumnIndex());
                    mergeCells(cellStartRange, cellEndRange, table);
                }
            }
            builder.endTable();
        }
        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), "testD", ".docx"));
    }


    @Test
    public void testF() throws Exception {
        List<String> stringList = new ArrayList<String>(12);
        for (int i = 0; i < 12; i++) {
            stringList.add(UUID.randomUUID().toString());
        }
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (String s : stringList) {
            builder.writeln(s);
            List<MergeCellModel> mergeCellModelList = Lists.newArrayList();
            Table table = builder.startTable();
            //行
            for (int j = 0; j < 20; j++) {
                if (0 <= j && j <= 5) {
                    //列
                    for (int k = 0; k < 5; k++) {
                        if (k == 0 && j == 0) {
                            builder.insertCell();
                            builder.writeln("区域位置");
                        } else {
                            builder.insertCell();
                            builder.writeln(String.format("%d-%d", j, k));
                        }
                    }
                    builder.endRow();
                    mergeCellModelList.add(new MergeCellModel(0, 0, 5, 0));
                    mergeCellModelList.add(new MergeCellModel(j, 1, j, 2));
                }
                if (5 < j && j <= 10) {
                    for (int k = 0; k < 5; k++) {
                        if (k == 0 && j == 6) {
                            builder.insertCell();
                            builder.writeln("交通状况");
                        } else {
                            builder.insertCell();
                            builder.writeln(String.format("%d-%d", j, k));
                        }
                    }
                    builder.endRow();
                    mergeCellModelList.add(new MergeCellModel(6, 0, 10, 0));
                    mergeCellModelList.add(new MergeCellModel(j, 1, j, 2));
                }
                if (10 < j && j <= 15) {
                    for (int k = 0; k < 5; k++) {
                        if (k == 0 && j == 11) {
                            builder.insertCell();
                            builder.writeln("外部配套设施");
                        } else {
                            builder.insertCell();
                            builder.writeln(String.format("%d-%d", j, k));
                        }
                    }
                    builder.endRow();
                    mergeCellModelList.add(new MergeCellModel(11, 0, 15, 0));
                    mergeCellModelList.add(new MergeCellModel(12, 1, 15, 1));
                    mergeCellModelList.add(new MergeCellModel(11, 1, 11, 2));
                }
                if (15 < j && j <= 18) {
                    for (int k = 0; k < 5; k++) {
                        if (k == 0 && j == 16) {
                            builder.insertCell();
                            builder.writeln("周围环境和景观");
                        } else {
                            builder.insertCell();
                            builder.writeln(String.format("%d-%d", j, k));
                        }
                    }
                    builder.endRow();
                    mergeCellModelList.add(new MergeCellModel(16, 0, 18, 0));
                    mergeCellModelList.add(new MergeCellModel(j, 1, j, 2));
                }
            }
            for (MergeCellModel mergeCellModel : mergeCellModelList) {
                Cell cellStartRange = table.getRows().get(mergeCellModel.getStartRowIndex()).getCells().get(mergeCellModel.getStartColumnIndex());
                Cell cellEndRange = table.getRows().get(mergeCellModel.getEndRowIndex()).getCells().get(mergeCellModel.getEndColumnIndex());
                mergeCells(cellStartRange, cellEndRange, table);
            }
            builder.endTable();
        }
        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), "testV", ".docx"));
    }

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
