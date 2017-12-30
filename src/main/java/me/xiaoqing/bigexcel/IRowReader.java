package me.xiaoqing.bigexcel;

import java.util.List;

/**
 * Created by za-dingxiaoqing on 2017/12/21.
 */
public class IRowReader implements IRowReaderInterface {

    public void getRows(int sheetIndex, int curRow, List<String> rowlist) {
        if(curRow % 1000 == 0){
            System.out.print(curRow+" ");
            for (int i = 0; i < rowlist.size(); i++) {
                System.out.print(rowlist.get(i) + " ");
            }
            System.out.println();
        }
    }
}
