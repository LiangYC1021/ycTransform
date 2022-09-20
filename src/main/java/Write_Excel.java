import java.io.*;
import java.nio.channels.FileChannel;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.usermodel.*;

import static java.lang.Math.min;

/**
 * 生成的 Excel 格式为 xlsx
 * 单例模式 只有一个Write_Excel类
 * 表头：
 * 日期	         需求名称	研发单名称  	配置说明	 开发者	启动应用   影响范围  修改说明
 * 文件最后修改日期                                 null    null      null    null
 */

public class Write_Excel {
    //总的工作簿（只有一个工作簿）
    public static XSSFWorkbook workbook;
    public static XSSFSheet con_sheet;
    public static XSSFSheet sql_sheet;
    public static XSSFSheet task_sheet;

    //研发单号-负责人名字 的map
    public static Map<String, String> responsible;
    //行高
    public static final short rowHeight = 1024;
    //要提取出来的大文件的路径列表 1.5版本已废弃
    public static List<String> big_files;
    //要生成的总文件夹的路径
    public static String firstDirPath;
    //要生成的总SQL文件夹的路径 包含在 firstDir里
    public static String sqlDirPath;
    //对于同名的sql文件，用这个map记录这个名字出现了多少次
    private static Map<String,Integer> nameCnt;
    public Write_Excel(String file_date) {
        // 创建工作薄
        workbook = new XSSFWorkbook();
        responsible = new HashMap<String, String>();
        //创建工作表
        create_Sheet();
        //初始化大文件列表
        big_files = new ArrayList<>();
        //初始化两个路径字符串 并生成两个文件夹

        //文件夹名字和excel的名字
        String name = "北京版本升级列表-" + file_date;
        //新建文件夹
        File directory = new File(name);
        if (!directory.exists()) directory.mkdir();
        //设置firstDir的路径
        firstDirPath=directory.getPath();
//        1.5 : 生成一个总的SQL文件夹 里面包含了n种sql类型的文件夹
//        SQL文件->mysql
//               ->oracle
//               ->postgresql...
        File sqlDir=new File(firstDirPath,"SQL文件");
        if(!sqlDir.exists())sqlDir.mkdir();
        sqlDirPath=sqlDir.getPath();

        //初始化map
        nameCnt=new HashMap<String,Integer>();
    }
    public Write_Excel(){}

    public void create_config_sheet() {
        // 创建工作表
        con_sheet = workbook.createSheet("配置");
        //创建表头
        XSSFRow conHeadRow = con_sheet.createRow(0);
        //设置表头格式和样式
        for (int i = 0; i < 8; i++) {
            String val = new String();
            conHeadRow.createCell(i).setCellStyle(cell_style(workbook, true));
            switch (i) {
                case 0:
                    val = "日期";
                    con_sheet.setColumnWidth(i, 10 * 256);
                    break;
                case 1:
                    val = "研发单号";
                    con_sheet.setColumnWidth(i, 20 * 256);
                    break;
                case 2:
                    val = "研发单名称";
                    con_sheet.setColumnWidth(i, 35 * 256);
                    break;
                case 3:
                    val = "配置说明";
                    con_sheet.setColumnWidth(i, 50 * 256);
                    break;
                case 4:
                    val = "开发者";
                    con_sheet.setColumnWidth(i, 10 * 256);
                    break;
                case 5:
                    val = "启动应用";
                    con_sheet.setColumnWidth(i, 30 * 256);
                    break;
                case 6:
                    val = "影响范围";
                    con_sheet.setColumnWidth(i, 25 * 256);
                    break;
                case 7:
                    val = "修改说明";
                    con_sheet.setColumnWidth(i, 50 * 256);
                    break;
                default:
            }
            conHeadRow.getCell(i).setCellValue(val);
        }
    }

    public void create_sql_sheet() {
        //创建工作表
        sql_sheet = workbook.createSheet("SQL");
        //创建表头
        XSSFRow sqlHeadRow = sql_sheet.createRow(0);
        //设置表头格式和样式
        for (int i = 0; i < 4; i++) {
            String val = new String();
            sqlHeadRow.createCell(i).setCellStyle(cell_style(workbook, true));
            switch (i) {
                case 0:
                    val = "日期";
                    sql_sheet.setColumnWidth(i, 10 * 256);
                    break;
                case 1:
                    val = "文件名称";
                    sql_sheet.setColumnWidth(i, 15 * 256);
                    break;
                case 2:
                    val = "SQL类型";
                    sql_sheet.setColumnWidth(i, 13 * 256);
                    break;
                case 3:
                    val = "文件内容";
                    sql_sheet.setColumnWidth(i, 90 * 256);
                    break;
                default:
            }
            sqlHeadRow.getCell(i).setCellValue(val);
        }
    }

//     任务类型	编号	标题	研发项目	负责人	状态	版本	应用模块	创建时间	创建人

    public void create_task_sheet(){
        //创建工作表
        task_sheet=workbook.createSheet("任务清单");
        //创建表头
        XSSFRow taskHeadRow = task_sheet.createRow(0);

        for(int i=0;i<22;i++){
            String val=new String();
            taskHeadRow.createCell(i).setCellStyle(cell_style(workbook, true));
            switch (i) {
                case 0:
                    val = "任务类型";task_sheet.setColumnWidth(i, 10 * 256);break;
                case 1:
                    val = "编号";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 2:
                    val = "标题";task_sheet.setColumnWidth(i, 13 * 256);break;
                case 3:
                    val = "研发项目";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 4:
                    val = "负责人";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 5:
                    val = "状态";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 6:
                    val = "版本";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 7:
                    val = "应用模块";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 8:
                    val = "创建时间";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 9:
                    val = "创建人";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 10:
                    val = "开始时间";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 11:
                    val = "截止时间";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 12:
                    val = "优先级";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 13:
                    val = "来源";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 14:
                    val = "分类";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 15:
                    val = "预估工时";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 16:
                    val = "所属需求";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 17:
                    val = "标签";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 18:
                    val = "故事点";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 19:
                    val = "设计负责人";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 20:
                    val = "开发负责人";task_sheet.setColumnWidth(i, 15 * 256);break;
                case 21:
                    val = "测试负责人";task_sheet.setColumnWidth(i, 15 * 256);break;
                default:
            }
            taskHeadRow.getCell(i).setCellValue(val);
        }
    }
    /**
     * 创建3个新的sheet
     */
    public void create_Sheet() {
        create_config_sheet();
        create_sql_sheet();
        create_task_sheet();
    }

    /**
     * 将word文档解析完的、已经配好格式的内容插入到配置sheet里
     *
     * @param list word文档解析完的、已经配好格式的内容
     */
    public void insert_config_sheet(List<List<String>> list) {
        //遍历数据并写入sheet
        for (int i = 0; i < list.size(); i++) {
            int row = con_sheet.getLastRowNum() + 1;
            XSSFRow sheetRow = con_sheet.createRow(row);
            List<String> data = list.get(i);
            for (int col = 0; col < data.size(); col++) {
                sheetRow.createCell(col).setCellValue(data.get(col));
                sheetRow.getCell(col).setCellStyle(cell_style(workbook, false));
                sheetRow.setHeight(rowHeight);
            }
        }
    }

    /**
     * 作用同 insert_config_sheet
     *
     * @param list
     */
    public void insert_sql_sheet(List<String> list) {
        int row = sql_sheet.getLastRowNum() + 1;
        XSSFRow sheetRow = sql_sheet.createRow(row);
        for (int col = 0; col < list.size(); col++) {
            if (list.get(col).length() > 30000) {
                //进行单元格拆分
                System.out.println("长度为：" + list.get(col).length());
                String content = list.get(col);
                sheetRow.createCell(col).setCellValue(content.substring(0, 30000));
                //新建空行
                int cur = 30000;
                while (cur < content.length()) {
                    List<String> tmp = new ArrayList<>();
                    //前三列为空
                    for (int i = 0; i < 3; i++) {
                        tmp.add("");
                    }
                    tmp.add(content.substring(cur, min(cur + 30000 - 1, content.length() - 1)));
                    insert_sql_sheet(tmp);
                    cur += 30000;
                }
            } else {
                sheetRow.createCell(col).setCellValue(list.get(col));
            }
            sheetRow.getCell(col).setCellStyle(cell_style(workbook, false));
            sheetRow.setHeight(rowHeight);
        }
    }


    /**
     * 与上面两个insert方法相同
     * @param list 提取出来的一整行元素
     */
    public void insert_task_sheet(List<HSSFRow> list){
        for(HSSFRow row:list){
            //新建一个行
            task_sheet.createRow(task_sheet.getLastRowNum()+1);
            //获取最新的行
            XSSFRow sheetRow= task_sheet.getRow(task_sheet.getLastRowNum());
            for(int i=0;i<row.getLastCellNum();i++){
                HSSFCell cell=row.getCell(i);
                sheetRow.createCell(i).setCellValue(cell.getStringCellValue());
            }
        }
    }

    /**
     * 将这个xlsx里的内容解析一下，生成任务编号对应开发负责人的map
     * row(1)     row(20)
     * 编号       开发负责人
     *
     * @param xlsx
     */
    public void set_hash_map(File xlsx) {
        try {
            HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(xlsx.getPath()));
            HSSFSheet sheet = wb.getSheetAt(0);
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                HSSFRow curRow = sheet.getRow(i);
                HSSFCell cell1 = curRow.getCell(1), cell20 = curRow.getCell(20);
                String id = cell1.toString();
                String name = cell20.toString();
                if (name.isEmpty()) {
                    HSSFCell cell19 = curRow.getCell(19);
                    name = cell19.toString();
                }
                if (name.isEmpty()) {
                    HSSFCell cell9 = curRow.getCell(9);
                    name = cell9.toString();
                }
                responsible.put(id, name);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }


    private static XSSFCellStyle cell_style(XSSFWorkbook workbook, boolean isHead) {
        XSSFCellStyle res = workbook.createCellStyle();
        //如果是表头要设置内容居中和背景色
        if (isHead) {
            //内容居中
            res.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            res.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            //设置填充的背景色
            res.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
            res.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            //设置字体加粗
            res.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        } else {
            //内容垂直居中
            res.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            //自动换行
            res.setWrapText(true);
        }
        // 边框颜色 黑色
        res.setTopBorderColor(IndexedColors.BLACK.getIndex());
        res.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        res.setRightBorderColor(IndexedColors.BLACK.getIndex());
        res.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        // 边框线型
        res.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
        res.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
        res.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
        res.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
        return res;
    }

    private String deal_same_name_sql_file(String name) {
        int cnt= nameCnt.get(name)==null ? 0:nameCnt.get(name);
        nameCnt.put(name,cnt+1);
        if(cnt==0)return name;
        else {
            String res=new String();
            for(int i=name.length()-1;i>=0;i--){
                if(name.charAt(i)=='.'){
                    res = name.substring(0,i)+"("+cnt+").sql";
                    break;
                }
            }
            return res;
        }
    }
    /**
     * 创建工作簿、扫描文档、解析文档内容、创建工作表、分别对两种工作表进行insert
     * 上述流程结束后，调用这个函数
     * 将完成的工作簿写入excel文件
     * @param file_date 要生成的文件夹的后缀的那个数字  即 北京升级列表-********
     */
    public void write_to_file(String file_date) {
        //文件夹名字和excel的名字
        String name = "北京版本升级列表-" + file_date;

        String path = firstDirPath;
        System.out.println("生成文件夹： "+path);
        //写入文件
        try {
            workbook.write(new FileOutputStream(new File(path, name + ".xlsx")));
            workbook.close();
            generate_big_files(path);
        } catch (Exception ex) {
            System.out.println("name");
            System.out.println(ex.getMessage());
        }
    }

    public static void insert_big_files(String path) {
        big_files.add(path);
        System.out.println("大文件：" + path);
    }

    // TODO: 2022/9/5 看看现场那边是否需要将每个sql文件放到文件夹里
    public static void generate_big_files(String directory_path) throws IOException {
        for (String path : big_files) {
            //获取源文件内容
            File source = new File(path);
            //同步源文件和要生成的文件的名字
            String name = source.getName();
            //获取父文件夹名字（sql类型），后续在对应sql文件名前加上sql类型
            File parent = source.getParentFile();

            File dest = new File(directory_path, parent.getName() + "-" + name);
            FileChannel sourceChannel = null;
            FileChannel destChannel = null;
            try {
                sourceChannel = new FileInputStream(source).getChannel();
                destChannel = new FileOutputStream(dest).getChannel();
                destChannel.transferFrom(sourceChannel, 0, sourceChannel.size());
            } finally {
                sourceChannel.close();
                destChannel.close();
            }

        }
    }

    /**
     * 生成sql文件到指定目录下，这个目录是它的sql类型
     * @param sqlType sql类型
     * @param sqlPath sql源文件的地址
     * @throws IOException
     */
    // TODO: 2022/9/20 处理 生成同名文件的问题，对于同名文件 在后面加（1）（2）这样子
    public void generate_sql_files(String sqlType,String sqlPath) throws IOException {
        //这个sql类型的文件夹 若不存在则创建文件夹
        File singleTypeDir=new File(sqlDirPath,sqlType);
        if(!singleTypeDir.exists())singleTypeDir.mkdir();
        //源文件
        File source=new File(sqlPath);
        //目标文件
        File dest=new File(singleTypeDir.getPath(),deal_same_name_sql_file(source.getName()));

        FileChannel sourceChannel = null;
        FileChannel destChannel = null;
        try {
            sourceChannel = new FileInputStream(source).getChannel();
            destChannel = new FileOutputStream(dest).getChannel();
            destChannel.transferFrom(sourceChannel, 0, sourceChannel.size());
        } finally {
            sourceChannel.close();
            destChannel.close();
        }
    }
}
