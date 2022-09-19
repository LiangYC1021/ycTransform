import java.io.IOException;
import java.text.DateFormat;
import java.util.*;
import java.io.File;
import com.spire.doc.*;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;


/**
 * 读取的 Word 格式为 docx
 * 这个类读取、操作的都是单个word文档
 * totalList中保存的是这个文档中所有的内容
 * 表头：
 * 日期	         研发单号	研发单名称  	配置说明	 开发者	启动应用   影响范围  修改说明
 * 文件最后修改日期                                 null    null      null    null
 */
public class Read_Word {
    /**
     * totalList 的内容：
     * 0:日期1  研发单号（同时也是研发单名称）1  配置说明1
     * 1:日期2  研发单号（同时也是研发单名称）2  配置说明2
     * ....
     * 由于该类操作的是一个单一的word文档，所以日期都是相同的
     */
    private List<List<String>> totalList = new LinkedList<List<String>>();
    private String date = new String();

    public List<List<String>> getTotalList() {
        return totalList;
    }


    public List<String> searchWordDocX(String fileUrl) {
        //获取文件日期
        File file = new File(fileUrl);
        DateFormat df = DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.CHINA);
        date = df.format(file.lastModified());
        //读取文件路径
        OPCPackage opcPackage = null;
        String content = null;
        List<String> docxList = new ArrayList<String>();
        try {
//            opcPackage = POIXMLDocument.openPackage(request.getSession().getServletContext().getRealPath(fileUrl));
            opcPackage = POIXMLDocument.openPackage(fileUrl);
            XWPFDocument xwpf = new XWPFDocument(opcPackage);
            POIXMLTextExtractor poiText = new XWPFWordExtractor(xwpf);
            content = poiText.getText();
//            System.out.println("content:\n"+content);
            StringBuffer cur = new StringBuffer();
            for (int i = 0; i < content.length(); i++) {
                char ch = content.charAt(i);
                cur.append(ch);
                if (ch == '\n') {
                    docxList.add(cur.toString());
                    cur = new StringBuffer();
                }
            }
//            docxList.add(content);
        } catch (IOException e) {
            e.printStackTrace();
        }
        docxList.add("#1");
        return docxList;
    }

    /**
     * 解析一个word文档的内容，把其中的内容不断地加到totalList里
     * res:包含了每个新行里的 配置说明 单元格内容
     *
     * @param list 该word文档里的所有内容 （一行一行地存储进这个list里）
     */
    public void parseSingleWord(List<String> list) {
        List<String> res = new ArrayList<>();
        boolean isNewRow[] = new boolean[list.size() + 1];
        String ID[] = new String[list.size() + 1];
        for (int i = 0; i < list.size(); i++) {
            String cur = list.get(i);
            if (cur.charAt(0) == '#' && '0' <= cur.charAt(1) && cur.charAt(1) <= '9') {
                isNewRow[i] = true;
                for (int j = 1; j < cur.length(); j++) {
                    if (!('0' <= cur.charAt(j) && cur.charAt(j) <= '9')) {
                        ID[i] = cur.substring(1, j);
//                        System.out.println("id = "+ID[i] + " " +ID[i].length());
                        break;
                    }
                }
            }
        }

        StringBuffer content = new StringBuffer();
        for (int i = 0; i < list.size(); i++) {
            String cur = list.get(i);
            if (isNewRow[i] == true) {
                if (content.length() > 0) {
                    while (content.charAt(0) == '\n') {
                        content.deleteCharAt(0);
                    }
                    while (content.charAt(content.length() - 1) == '\n') {
                        content.deleteCharAt(content.length() - 1);
                    }
                    res.add(content.toString());
//                    System.out.println("这一段的内容为\n---------"+content+"----------");
                    content = new StringBuffer("");
                }
            } else {
                content.append(cur);
            }
        }

        int id = 0;
        for (int i = 0; i < list.size() - 1; i++) {
            if (isNewRow[i]) {
                String con=res.get(id++);
                if(con.length()>30000){
                    con=new String("文件长度为："+con.length()+"已超出最大长度限制\n具体内容详见附件");
                }
                List<String> tmp = add_to_totalList(date, list.get(i), con, Write_Excel.responsible.get(ID[i]));
                totalList.add(tmp);
            }
        }
    }

    private static List<String> add_to_totalList(String dat, String nam, String con, String developer) {
//         * 日期	         研发单号	研发单名称  	配置说明	 开发者	启动应用   影响范围  修改说明
//         * 文件最后修改日期                                          null      null    null
        List<String> tmp = new ArrayList<>();
        nam = nam.substring(0, nam.length() - 1);//去除行末的换行符
        //获取研发单号
        String number = "";
        for (int i = 1; i < nam.length(); i++) {
            char ch = nam.charAt(i);
            if ('0' <= ch && ch <= '9') {
                number += ch;
            } else break;
        }
        tmp.add(dat);
        tmp.add(number);
        tmp.add(nam);
        tmp.add(con);
        tmp.add(developer);
        tmp.add("");
        tmp.add("");
        tmp.add("");
        return tmp;
    }
}
