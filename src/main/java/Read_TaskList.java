import java.io.*;
import java.text.DateFormat;
import java.util.*;

import com.spire.doc.*;

import java.io.IOException;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * 读取任务清单 格式为 xlsx
 * 这个类读取、操作的都是单个excel文档
 * 表头：
 * 任务类型	编号	标题	研发项目	负责人	状态	版本	应用模块	创建时间	创建人	开始时间	截止时间	优先级	来源	分类	预估工时	所属需求	标签	故事点	设计负责人	开发负责人	测试负责人
 */
public class Read_TaskList {
    public List<HSSFRow> totalList=new ArrayList<>();
    
    public void searchExcelXlsx(String fileUrl) throws IOException {
        InputStream is=new FileInputStream(fileUrl);
        //打开一个工作簿
        HSSFWorkbook excel=new HSSFWorkbook(is);
        //找到sheet1这个工作表
        HSSFSheet task_sheet=excel.getSheet("sheet1");
        //将每一行添加到totalList里面
        for(int i=1;i<=task_sheet.getLastRowNum();i++){
            HSSFRow row= task_sheet.getRow(i);
            totalList.add(row);
        }
    }

    public List<HSSFRow> getTotalList() {
        return totalList;
    }
}
