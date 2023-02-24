import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.json.JSONObject;

import java.io.*;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;

public class DuDoc4 {

    //    1、遍历文件夹下所有文件，返回文件名；
    static ArrayList<File> getFileAll(File file,ArrayList<File> fileList){

        File[] files = file.listFiles();
        for (int i = 0; i < files.length; i++) {
            //将文件添加到集合中
            if (!files[i].isHidden()){      //去除mac隐藏文件.dc_stoc
                fileList.add(files[i]);
            }
        }
        //返回所有文件名
        return fileList;
    }




    //    2、从各个文件中集读取数据，数组返回每个文件里的数据
    private static ArrayList readerr(ArrayList onn) {

        String st = "";
        ArrayList listt = new ArrayList();

        for (int i = 0;i < onn.size();i++){
            st = "";        //字符过多导致数据丢失，每次循环前重置字符串值
            File fo = new File(onn.get(i).toString());      //文件类

//            判断是不是文件，不是就不处理
            if (fo.exists()) {
                try {
                    //读取文件内容，装起来
                    FileInputStream file_in = new FileInputStream(fo);
                    InputStreamReader reader = new InputStreamReader(file_in, "UTF-8");
                    BufferedReader bufreader = new BufferedReader(reader);
                    String line;
                    while ((line = bufreader.readLine()) != null) {
                        st = st + line;
                    }
                    bufreader.close();
                    reader.close();
                    file_in.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (UnsupportedEncodingException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            listt.add(st);
        }
        return listt;
    }





    //    3、处理数据，二维数组返回需要的数据
    static ArrayList chuli(ArrayList listt){

        ArrayList l_all = new ArrayList();
        ArrayList<String> l_name = new ArrayList<String>();
        ArrayList<String> l_severity = new ArrayList<String>();
        ArrayList<String> l_issue_reporter = new ArrayList<String>();
        ArrayList<String> l_issue_operator = new ArrayList<String>();
        ArrayList<String> l_created_at = new ArrayList<String>();
        ArrayList<String> l_work_item_id = new ArrayList<String>();

        for (int i = 0;i < listt.size();i++){
            try{
                JSONObject json_test = new  JSONObject(listt.get(i).toString());

                int chang = json_test.getJSONObject("data").getJSONArray("work_items").length();
                for (int j = 0;j < chang;j++){
                    JSONObject sss = new JSONObject(json_test.getJSONObject("data").getJSONArray("work_items").get(j).toString());

//                    根据bug id去重
                    if(!l_work_item_id.contains("https://meego.feishu.cn/tiktok_live/issue/detail/"+sss.get("work_item_id").toString())){

//                        拿bug单标题
                        l_name.add(sss.getString("name"));

//                        拿严重程度
                        String severity_string = new String(sss.get("severity").toString());
                        int severity_int = Integer.valueOf(severity_string)-1;
                        l_severity.add("P"+severity_int);

//                        获取报告人
                        String issue_reporter_name = new JSONObject(sss.getJSONArray("issue_reporter").get(0).toString()).getString("nickname");
                        l_issue_reporter.add(issue_reporter_name);

//                        获取经办人
                        String issue_operator_name = new JSONObject(sss.getJSONArray("issue_operator").get(0).toString()).getString("nickname");
                        l_issue_operator.add(issue_operator_name);

//                        获取提单时间
                        String created_at_name = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date(Long.parseLong(String.valueOf(sss.get("created_at")))));
                        l_created_at.add(created_at_name);

//                        获取bug单链接
                        String bug_url = "https://meego.feishu.cn/tiktok_live/issue/detail/"+sss.get("work_item_id").toString();
                        l_work_item_id.add(bug_url);
                    }
                }

            }catch (Exception e){}
        }
        l_all.add(l_name);
        l_all.add(l_severity);
        l_all .add(l_issue_reporter);
        l_all.add(l_issue_operator);
        l_all.add(l_created_at);
        l_all.add(l_work_item_id);

        return l_all;
    }




    //    4、创建表格填入数据，导出文件
    static void creat_xls(ArrayList data_list) throws IOException, WriteException {

//        获取当前时间生成表格文件名字，放到特定的文件夹下
        LocalDateTime now = LocalDateTime.now();
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd HH.mm.ss");
        String xls_path = "/Users/bytedance/Desktop/工作/处理成的表格文件/"+dtf.format(now)+".xls";

//            jxl生成表格文件
        WritableWorkbook workBook = Workbook.createWorkbook(new File(xls_path));
        try{
//            jxl创建表1，填入每列每行对应的数据
            WritableSheet sheet = workBook.createSheet("工作表1", 0);

            for (int i = 0;i < data_list.size();i++){
                for (int j = 0;j < ((ArrayList)data_list.get(i)).size();j++){
                    switch (i){
                        case 0:     //填入第1列，bug标题
                            Label name = new Label(0,j,((ArrayList)data_list.get(i)).get(j).toString());
                            sheet.addCell(name);
                            break;
                        case 1:     //填入第2列，严重程度
                            Label severity = new Label(1,j,((ArrayList)data_list.get(i)).get(j).toString());
                            sheet.addCell(severity);
                            break;
                        case 2:     //填入第3列，提报人
                            Label issue_reporter = new Label(2,j,((ArrayList)data_list.get(i)).get(j).toString());
                            sheet.addCell(issue_reporter);
                            break;
                        case 3:     //填入第4列，经办人
                            Label issue_operator = new Label(3,j,((ArrayList)data_list.get(i)).get(j).toString());
                            sheet.addCell(issue_operator);
                            break;
                        case 4:     //填入第5列，提出时间
                            Label created_at = new Label(4,j,((ArrayList)data_list.get(i)).get(j).toString());
                            sheet.addCell(created_at);
                            break;
                        case 5:     //填入第6列，bug链接
                            Label work_item_id = new Label(5,j,((ArrayList)data_list.get(i)).get(j).toString());
                            sheet.addCell(work_item_id);
                            break;
                    }
                }
            }
        }catch (Exception e){}finally {
            workBook.write();
            workBook.close();       //疑问，为什么这样就不会报错？，在try里面为啥报错？
        }
    }




    public static void main(String[] args) throws IOException {
//        1、给定一个文件夹位置
        File file = new File("/Users/bytedance/Desktop/工作/抓取meego-组件化配置问题bug单数据");

//        2、调用方法获得文件夹下所有文件,装进动态数组里
        ArrayList path_list = new ArrayList<>();
        for (int i = 0;i < getFileAll(file,new ArrayList<File>()).size();i++){
            String path = getFileAll(file,new ArrayList<File>()).get(i).toString();
            path_list.add(path);
        }

//        3、测试调用：先从文件读取数据，然后交给数据处理，拿到需要的数据，然后创建表格文件填写进去，导出xls文件
        try {
            creat_xls(chuli(readerr(path_list)));
        }catch (Exception e){}

    }
}