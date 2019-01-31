package org.test;

import com.aspose.words.*;
import com.google.common.collect.Sets;
import com.help.TestFile;
import com.red.tool.other.MergeCellModel;
import org.apache.commons.collections.CollectionUtils;
import org.testng.annotations.Test;

import java.awt.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.UUID;

/**
 * @Auther: zch
 * @Date: 2019/1/23 15:11
 * @Description:
 */
public class Demo3Table {

    @Test
    public void testB() throws Exception {
        List<String> stringList = new ArrayList<String>(12);
        for (int i = 0; i < 12; i++) {
            stringList.add(UUID.randomUUID().toString().substring(0, 4));
        }
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("估价结果汇总表");
        Set<MergeCellModel> mergeCellModelList = Sets.newHashSet();
        Table table = builder.startTable();
        for (int j = 0; j < 2; j++) {
            switch (j) {
                case 0:
                    for (int k = 0; k < 8; k++) {
                        switch (k) {
                            case 0:
                                builder.insertCell();
                                builder.writeln("估价对象及结果\\估价方法及结果");
                                mergeCellModelList.add(new MergeCellModel(0, 0, 1, 2));
                                break;
                            case 3:
                                builder.insertCell();
                                builder.writeln("测算结果");
                                mergeCellModelList.add(new MergeCellModel(0, 3, 0, 5));
                                break;
                            case 6:
                                builder.insertCell();
                                builder.writeln("估价结果");
                                //未处理
                                mergeCellModelList.add(new MergeCellModel(j, 6, j+1, 7));
                                break;
                            default:
                                builder.insertCell();
                                break;
                        }
                    }
                    builder.endRow();
                    break;
                case 1:
                    for (int k = 0; k < 8; k++) {
                        switch (k) {
                            case 3:
                                builder.insertCell();
                                builder.writeln("市场比较法");
                                break;
                            case 4:
                                builder.insertCell();
                                builder.writeln("收益法");
                                break;
                            case 5:
                                builder.insertCell();
                                builder.writeln("假设开发法");
                                break;
                            case 6:
                                builder.insertCell();
                                mergeCellModelList.add(new MergeCellModel(j, 6, j, 7));
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
        }
        for (int j = 2; j < 2 + stringList.size(); j++) {
            int num = j % 2;
            switch (num) {
                case 0:
                    for (int k = 0; k < 8; k++) {
                        switch (k) {
                            case 0:
                                builder.insertCell();
                                builder.writeln(String.format("委估对象%s", stringList.get(j - 2)));
                                mergeCellModelList.add(new MergeCellModel(j, 0, j + 1, 0));
                                break;
                            case 1:
                                builder.insertCell();
                                builder.writeln("总价(元或万元)");
                                mergeCellModelList.add(new MergeCellModel(j, 1, j, 2));
                                break;
                            case 6:
                                builder.insertCell();
                                mergeCellModelList.add(new MergeCellModel(j, 6, j, 7));
                                break;
                            default:
                                builder.insertCell();
                                break;
                        }
                    }
                    builder.endRow();
                    System.out.println("0 ==> j:" + j);
                    break;
                case 1:
                    for (int k = 0; k < 8; k++) {
                        switch (k) {
                            case 1:
                                builder.insertCell();
                                builder.writeln("单价(元/m2)");
                                mergeCellModelList.add(new MergeCellModel(j, 1, j, 2));
                                break;
                            case 6:
                                builder.insertCell();
                                mergeCellModelList.add(new MergeCellModel(j, 6, j, 7));
                                break;
                            default:
                                builder.insertCell();
                                break;
                        }
                    }
                    builder.endRow();
                    System.out.println("1 ==> j:" + j);
                    break;
                default:
                    break;
            }
        }
        for (int j = 2 + stringList.size(); j < 2 + stringList.size() + 2; j++) {
            int num = j % 2;
            switch (num) {
                case 0:
                    for (int k = 0; k < 8; k++) {
                        switch (k) {
                            case 0:
                                builder.insertCell();
                                builder.writeln("汇总平均价值");
                                mergeCellModelList.add(new MergeCellModel(j, 0, j + 1, 0));
                                break;
                            case 1:
                                builder.insertCell();
                                builder.writeln("总价(元或万元)");
                                mergeCellModelList.add(new MergeCellModel(j, 1, j, 2));
                                break;
                            case 6:
                                builder.insertCell();
                                mergeCellModelList.add(new MergeCellModel(j, 6, j, 7));
                                break;
                            default:
                                builder.insertCell();
                                break;
                        }
                    }
                    builder.endRow();
                    System.out.println("0 ==> j:" + j);
                    break;
                case 1:
                    for (int k = 0; k < 8; k++) {
                        switch (k) {
                            case 1:
                                builder.insertCell();
                                builder.writeln("平均单价(元/m2)");
                                mergeCellModelList.add(new MergeCellModel(j, 1, j, 2));
                                break;
                            case 6:
                                builder.insertCell();
                                mergeCellModelList.add(new MergeCellModel(j, 6, j, 7));
                                break;
                            default:
                                builder.insertCell();
                                break;
                        }
                    }
                    builder.endRow();
                    System.out.println("1 ==> j:" + j);
                    break;
                default:
                    break;
            }
        }
        if (CollectionUtils.isNotEmpty(mergeCellModelList)) {
            for (MergeCellModel mergeCellModel : mergeCellModelList) {
                Cell cellStartRange = table.getRows().get(mergeCellModel.getStartRowIndex()).getCells().get(mergeCellModel.getStartColumnIndex());
                Cell cellEndRange = table.getRows().get(mergeCellModel.getEndRowIndex()).getCells().get(mergeCellModel.getEndColumnIndex());
                mergeCells(cellStartRange, cellEndRange, table);
            }
        }
        builder.endTable();
        //当前类 和 当前运行方法
        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), Thread.currentThread().getStackTrace()[1].getMethodName(), ".docx"));
    }

    @Test
    public void testA() throws Exception {
        List<String> stringList = new ArrayList<String>(12);
        for (int i = 0; i < 12; i++) {
            stringList.add(UUID.randomUUID().toString().substring(0, 4));
        }
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("估价结果一览表");
        int i = 0;
        while (i < stringList.size() + 1) {
            Set<MergeCellModel> mergeCellModelList = Sets.newHashSet();
            Table table = builder.startTable();
            switch (i) {
                case 1:
                    for (int k = 0; k < 9; k++) {
                        switch (k) {
                            case 0:
                                builder.insertCell();
                                builder.writeln("产权证号");
                                break;
                            case 1:
                                builder.insertCell();
                                builder.writeln("业务件号");
                                break;
                            case 2:
                                builder.insertCell();
                                builder.writeln("房屋座落");
                                break;
                            case 3:
                                builder.insertCell();
                                builder.writeln("权利人");
                                break;
                            case 4:
                                builder.insertCell();
                                builder.writeln("共有方式");
                                break;
                            case 5:
                                builder.insertCell();
                                builder.writeln("规划用途");
                                break;
                            case 6:
                                builder.insertCell();
                                builder.writeln("所在层数");
                                break;
                            case 7:
                                builder.insertCell();
                                builder.writeln("建筑面积");
                                break;
                            case 8:
                                builder.insertCell();
                                builder.writeln("房地产单价");
                                break;
                            case 9:
                                builder.insertCell();
                                builder.writeln("房地产价值");
                                break;
                            default:
                                builder.insertCell();
                                break;
                        }
                    }
                    builder.endRow();
                    break;
                case 2:
                    this.testA2(builder, stringList.get(i - 1));
                    break;
                case 3:
                    this.testA2(builder, stringList.get(i - 1));
                    break;
                default:
                    break;
            }

            i++;

            builder.getCellFormat().getBorders().getLeft().setLineWidth(1.0);
            builder.getCellFormat().getBorders().getRight().setLineWidth(1.0);
            builder.getCellFormat().getBorders().getTop().setLineWidth(1.0);
            builder.getCellFormat().getBorders().getBottom().setLineWidth(1.0);
            builder.getCellFormat().setWidth(100);
            builder.getCellFormat().setVerticalMerge(CellVerticalAlignment.CENTER);
            builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
            builder.endTable();
        }

        //当前类 和 当前运行方法
        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), Thread.currentThread().getStackTrace()[1].getMethodName(), ".docx"));
    }


    private void testA2(DocumentBuilder builder, String str) throws Exception {
        for (int k = 0; k < 9; k++) {
            switch (k) {
                case 0:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 1:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 2:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 3:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 4:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 5:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 6:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 7:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 8:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 9:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                default:
                    builder.insertCell();
                    break;
            }
        }
        builder.endRow();
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
