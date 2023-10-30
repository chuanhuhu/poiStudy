package com.hyc.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.PageReadListener;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.fastjson2.JSON;
import com.hyc.easyexcel.mode.entity.Students;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@Slf4j
@SpringBootTest
class EasyexcelApplicationTests {

    /**
     * 最简单的读
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link Students}
     * <p>
     * 2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link StudentsListener}
     * <p>
     * 3. 直接读即可
     */
//    @Test
//    public void simpleRead() {
//        // 写法1：JDK8+ ,不用额外写一个StudentsListener
//        // since: 3.0.0-beta1
//        String fileName = TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
//        // 这里默认每次会读取100条数据 然后返回过来 直接调用使用数据就行
//        // 具体需要返回多少行可以在`PageReadListener`的构造函数设置
//        EasyExcel.read(fileName, Students.class, new PageReadListener<Students>(dataList -> {
//            for (Students demoData : dataList) {
//                log.info("读取到一条数据{}", JSON.toJSONString(demoData));
//            }
//        })).sheet().doRead();
//
//        // 写法2：
//        // 匿名内部类 不用额外写一个StudentsListener
//        fileName = TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
//        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
//        EasyExcel.read(fileName, Students.class, new ReadListener<Students>() {
//            /**
//             * 单次缓存的数据量
//             */
//            public static final int BATCH_COUNT = 100;
//            /**
//             *临时存储
//             */
//            private List<Students> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
//
//            @Override
//            public void invoke(Students data, AnalysisContext context) {
//                cachedDataList.add(data);
//                if (cachedDataList.size() >= BATCH_COUNT) {
//                    saveData();
//                    // 存储完成清理 list
//                    cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
//                }
//            }
//
//            @Override
//            public void doAfterAllAnalysed(AnalysisContext context) {
//                saveData();
//            }
//
//            /**
//             * 加上存储数据库
//             */
//            private void saveData() {
//                log.info("{}条数据，开始存储数据库！", cachedDataList.size());
//                log.info("存储数据库成功！");
//            }
//        }).sheet().doRead();
//
//        // 有个很重要的点 StudentsListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
//        // 写法3：
//        fileName = TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
//        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
//        EasyExcel.read(fileName, Students.class, new StudentsListener()).sheet().doRead();
//
//        // 写法4
//        fileName = TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
//        // 一个文件一个reader
//        try (ExcelReader excelReader = EasyExcel.read(fileName, Students.class, new StudentsListener()).build()) {
//            // 构建一个sheet 这里可以指定名字或者no
//            ReadSheet readSheet = EasyExcel.readSheet(0).build();
//            // 读取一个sheet
//            excelReader.read(readSheet);
//        }
//    }


    public static void main(String[] args) {
        try {
            File file = ResourceUtils.getFile("classpath:student.xlsx");
            EasyExcel.read(file)
                    .excelType(ExcelTypeEnum.XLSX)
                    .sheet();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }
}
