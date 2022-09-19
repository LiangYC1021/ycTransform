import java.io.*;
import java.text.DateFormat;
import java.util.*;


/**
 * 读取的 SQL
 * 这个类读取、操作的都是单个 SQL文件
 * totalList中保存的是这个文件中所有的内容
 * 表头：
 * 日期	  SQL文件名称  SQL类型  SQL文件内容
 */

public class Read_SQL {
    /**
     * totalList 的内容：
     * 0:日期1  SQL文件名称1  SQL类型1  SQL文件内容1
     * 只有一行
     * 由于该类操作的是一个单一的 SQL文件，所以日期都是相同的
     */
    private List<String> totalList = new ArrayList<>();

    public List<String> getTotalList() {
        return totalList;
    }

    public void parseSQLFile(String sqlType, String fileUrl) {
        //表元素
        String date;
        String fileName;
        String content;

        File sqlFile = new File(fileUrl);
        //获取文件日期
        File file = new File(fileUrl);
        DateFormat df = DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.CHINA);
        date = df.format(file.lastModified());
        fileName = sqlFile.getName();
        //获取文件内容
        content = readFileContent(fileUrl);
        //添加至totalList
        totalList.add(date);
        totalList.add(fileName);
        totalList.add(sqlType);
        if (content.length() <= 30000) {
            totalList.add(content);
        } else {
            String tmp=sqlFile.getParentFile().getName()+"-"+fileName;
            totalList.add("文件长度为： " + content.length() + " 已超出最大长度限制\n具体内容详见附件:"+tmp);
            Write_Excel.insert_big_files(fileUrl);
        }

//        输出测试
//        System.out.println(date+" "+fileName+"  type = "+sqlType);
//        System.out.println(content);
    }

    public static String readFileContent(String fileUrl) {
        File file = new File(fileUrl);
        BufferedReader reader = null;
        StringBuffer sbf = new StringBuffer("#\n");
        FileInputStream fis = null;
        InputStreamReader read = null;
        try {
            fis = new FileInputStream(file);
            read = new InputStreamReader(fis, "UTF-8");
            reader = new BufferedReader(read);
            String tempStr;
            while ((tempStr = reader.readLine()) != null) {
                sbf.append(tempStr);
                sbf.append('\n');
            }

            reader.close();
            return sbf.toString();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (reader != null) {
                try {
                    reader.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        }
        return sbf.toString();
    }

}
