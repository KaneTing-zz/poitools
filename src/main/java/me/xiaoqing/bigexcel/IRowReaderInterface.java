package me.xiaoqing.bigexcel;

import java.util.List;

/**
 * Created by za-dingxiaoqing on 2017/12/21.
 */
public interface IRowReaderInterface {

    /**
     * 业务逻辑实现方法
     * @param sheetIndex
     * @param curRow
     * @param rowlist
     */
    void getRows(int sheetIndex, int curRow, List<String> rowlist);
}
