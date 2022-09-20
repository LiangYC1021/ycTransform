import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * @author Liang Yucheng
 * @date 2022/8/22
 * Steps: 创建工作簿、扫描文档、解析文档内容、创建工作表、分别对两种工作表进行insert
 */

public class APP {
    public static Write_Excel excel = new Write_Excel();//只有一个excel对象
    //压缩包路径 后缀固定为 .zip
    public static String zip_path = new String();

    public static void main(String[] args) {
        new AppFrame();
    }

    public static void Run() {

        //压缩包解压后的文件夹名
        String path = new String(zip_path.substring(0, zip_path.length() - 4));
        try {
            unZipFiles(zip_path, path);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        File file = new File(path);        //获取其file对象
        String file_date = new String();
        for (int i = path.length() - 1; i >= 0; i--) {
            if (path.charAt(i) == '-') {
                file_date = path.substring(i + 1);
                break;
            }
        }
        excel = new Write_Excel(file_date);
        File[] fs = file.listFiles();    //遍历path下的文件和目录，放在File数组中
        for (File dir : fs) {                    //遍历File[]数组
            enter_directory(dir);
        }

        excel.write_to_file(file_date);
        System.out.println("Excel generated successfully!");
    }

    /**
     * 进入了一个文件夹 然后需要在这里插入到两个sheet里
     *
     * @param dir
     */
    public static void enter_directory(File dir) {
        File[] fs = dir.listFiles();
        for (File f : fs) {
            if ("配置说明".equals(f.getName())) {
                File[] config = f.listFiles();
                for (File con : config) {
                    if (con.isDirectory()) continue;
                    //创建 解析当前这个word配置文档（已到最小单位）
                    Read_Word read_word = new Read_Word();
                    //生成totalList
                    read_word.parseSingleWord(read_word.searchWordDocX(con.getPath()));
                    //插入该文档的内容到配置sheet里
                    excel.insert_config_sheet(read_word.getTotalList());
                }
            } else if ("SQL脚本".equals(f.getName())) {
                //此时对应的是SQL的类型的文件夹
                for (File sqlType : f.listFiles()) {
                    String typeName = sqlType.getName();
                    //此时进入了这个类型的文件夹 比如说MySQL文件夹
                    for (File sql : sqlType.listFiles()) {
                        if(sql.isDirectory())continue;
                        //已到最小单位 SQL文件
                        Read_SQL read_sql = new Read_SQL();
                        read_sql.parseSQLFile(typeName, sql.getPath());
                        excel.insert_sql_sheet(read_sql.getTotalList());

                        //1.5版本 生成对应的sql文件附件
                        try {
                            excel.generate_sql_files(typeName,sql.getPath());
                        } catch (IOException e) {
                            throw new RuntimeException(e);
                        }

                    }
                }
            } else if ("任务清单".equals(f.getName())) {
                for (File xlsx : f.listFiles()) {
                    excel.set_hash_map(xlsx);
                    Read_TaskList read_taskList=new Read_TaskList();
                    try {
                        read_taskList.searchExcelXlsx(xlsx.getPath());
                        excel.insert_task_sheet(read_taskList.getTotalList());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                }
            }
        }
    }

    /**
     * zip文件解压
     *
     * @param inputFile   待解压文件夹/文件
     * @param destDirPath 解压路径
     */
    public static void unZipFiles(String inputFile, String destDirPath) throws Exception {
        File srcFile = new File(inputFile);//获取当前压缩文件
        // 判断源文件是否存在
        if (!srcFile.exists()) {
            throw new Exception(srcFile.getPath() + "所指文件不存在");
        }
        File destDir = new File(destDirPath);
        if (destDir.exists()) {
            destDir.delete();
        }
        ZipFile zipFile = new ZipFile(srcFile, Charset.forName("UTF-8"));//创建压缩文件对象
        //开始解压
        Enumeration<?> entries = zipFile.entries();
        while (entries.hasMoreElements()) {
            ZipEntry entry = (ZipEntry) entries.nextElement();
            // 如果是文件夹，就创建个文件夹
            if (entry.isDirectory()) {
                String dirPath = destDirPath + "/" + entry.getName();
                srcFile.mkdirs();
            } else {
                // 如果是文件，就先创建一个文件，然后用io流把内容copy过去
                File targetFile = new File(destDirPath + "/" + entry.getName());
                // 保证这个文件的父文件夹必须要存在
                if (!targetFile.getParentFile().exists()) {
                    targetFile.getParentFile().mkdirs();
                }
                targetFile.createNewFile();
                // 将压缩文件内容写入到这个文件中
                InputStream is = zipFile.getInputStream(entry);
                FileOutputStream fos = new FileOutputStream(targetFile);
                int len;
                byte[] buf = new byte[1024];
                while ((len = is.read(buf)) != -1) {
                    fos.write(buf, 0, len);
                }
                // 关流顺序，先打开的后关闭
                fos.close();
                is.close();
            }
        }
    }

}
