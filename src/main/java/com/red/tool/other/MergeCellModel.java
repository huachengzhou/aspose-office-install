package com.red.tool.other;

/**
 * @Auther: zch
 * @Date: 2019/1/18 19:53
 * @Description:
 */
public class MergeCellModel {

    private Integer startRowIndex;

    private Integer startColumnIndex;

    private Integer endRowIndex;

    private Integer endColumnIndex;


    public MergeCellModel(Integer startRowIndex, Integer startColumnIndex, Integer endRowIndex, Integer endColumnIndex) {
        this.startRowIndex = startRowIndex;
        this.startColumnIndex = startColumnIndex;
        this.endRowIndex = endRowIndex;
        this.endColumnIndex = endColumnIndex;
    }

    public Integer getStartRowIndex() {
        return startRowIndex;
    }

    public Integer getStartColumnIndex() {
        return startColumnIndex;
    }

    public Integer getEndRowIndex() {
        return endRowIndex;
    }

    public Integer getEndColumnIndex() {
        return endColumnIndex;
    }
}
