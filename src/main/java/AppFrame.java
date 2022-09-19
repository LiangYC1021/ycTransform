import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.awt.image.BufferedImage;

public class AppFrame extends Frame implements Runnable {
    /*****************************窗口相关属性**********************************/
    //刷新率
    public static final int REPAINT_INTERVAL = 30;
    //窗口宽高
    public static final int FRAME_WIDTH = 500;
    public static final int FRAME_HEIGHT = 350;
    //获得用户屏幕的宽高
    public static final int SCREEN_W = Toolkit.getDefaultToolkit().getScreenSize().width;
    public static final int SCREEN_H = Toolkit.getDefaultToolkit().getScreenSize().height;
    //窗口位置
    public static final int FRAME_X = (SCREEN_W - FRAME_WIDTH) >> 1;
    public static final int FRAME_Y = (SCREEN_H - FRAME_HEIGHT) >> 1;
    //窗口标题
    public static final String APP_TITLE = "ycTransform";
    //标题栏高度
    public static int titleBarH;
    /*****************************窗口相关属性**********************************/

//    private Image overImg=null;
//    //1.定义一张和屏幕大小一样的图片
//    private BufferedImage bufImg=new BufferedImage(FRAME_WIDTH,FRAME_HEIGHT,BufferedImage.TYPE_4BYTE_ABGR);
    //文本框 放路径
    JTextField filePath = new JTextField(40);
    //按钮
    Button open = new Button("Open");
    Button generate = new Button("Generate Excel");

    public AppFrame() {
        initFrame();
        initEventListener();
    }

    private void initFrame() {
        //取消布局
        setLayout(null);
        //设置标题
        setTitle(APP_TITLE);
        //设置窗口大小
        setSize(FRAME_WIDTH, FRAME_HEIGHT);
        //设置窗口的左上角的坐标
        setLocation(FRAME_X, FRAME_Y);
        //设置窗口大小不可改变
        setResizable(false);
        //设置窗口可见
        setVisible(true);
        titleBarH = getInsets().top;
        //文件路径文本框属性
        filePath.setSize(300, 50);
        filePath.setLocation(80, 50 + titleBarH);
        filePath.setEditable(false);
        filePath.setHorizontalAlignment(JTextField.CENTER);
        filePath.setFont(new Font("宋体", Font.BOLD, 13));
        add(filePath);
        //选择按钮属性
        open.setSize(50, 50);
        open.setLocation(80 + filePath.getWidth() + 35, 50 + titleBarH);
        add(open);
        //确定按钮属性
        generate.setSize(100, 50);
        generate.setLocation(200, 200 + titleBarH);
        add(generate);
    }

    @Override
    public void run() {
        while (true) {
            //在此调用repaint，回调update
            repaint();
            try {
                Thread.sleep(REPAINT_INTERVAL);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }
//    public void update(Graphics g1){
//        //2.得到图片的画笔
//        Graphics g=bufImg.getGraphics();
//        //3.使用系统画笔，将图片绘制到frame上来
//        g1.drawImage(bufImg,0,0,null);
//    }

    private void initEventListener() {
        //注册监听事件
        addWindowListener(new WindowAdapter() {
            //点击关闭按钮的时候，这个方法会被自动调用
            @Override
            public void windowClosing(WindowEvent e) {
                System.exit(0);
            }
        });

        open.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                //按钮点击事件
                JFileChooser chooser = new JFileChooser();             //设置选择器
                chooser.setMultiSelectionEnabled(false);             //设为多选
                int returnVal = chooser.showOpenDialog(open);        //是否打开文件选择框
                System.out.println("returnVal=" + returnVal);
                if (returnVal == JFileChooser.APPROVE_OPTION) {          //如果符合文件类型
                    String filepath = chooser.getSelectedFile().getAbsolutePath();      //获取绝对路径
                    filePath.setText(filepath);
//                    System.out.println("Absolute Path: "+filepath);
//                    System.out.println("You chose to open this file: " + chooser.getSelectedFile().getName());  //输出相对路径
                }
            }
        });

        generate.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String tmp = filePath.getText();
                if (!tmp.isEmpty()) {
                    APP.zip_path = new String(tmp);
                    filePath.setText("");
                    APP.Run();
                }
            }
        });
    }
}
