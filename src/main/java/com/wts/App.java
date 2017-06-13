package com.wts;

import com.wts.util.ParamesAPI;
import com.wts.util.PinyinTool;
import com.wts.util.WxService;
import me.chanjar.weixin.common.exception.WxErrorException;
import me.chanjar.weixin.cp.bean.WxCpMessage;
import me.chanjar.weixin.cp.bean.WxCpUser;
import net.sourceforge.pinyin4j.format.exception.BadHanyuPinyinOutputFormatCombination;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class App {

    public static void main(String[] args) throws IOException, BadHanyuPinyinOutputFormatCombination {
        System.out.println("                        欢迎使用槐荫人社信息发送简易程序");
        System.out.println(" ");
        System.out.println("------------------------------------用前须知------------------------------------");
        System.out.println(" ");
        System.out.println("1：本程序必须在办公外网环境下运行！");
        System.out.println("2：待查的Excel文件必须放置在C盘根目录下！");
        System.out.println("3：Excel文件后缀为xlsx，不支持xls后缀的低版本Excel文件！");
        System.out.println("4：Excel文件第一行为标题，列数视情况而定！");
        System.out.println("5：Excel文件第一列内容必须为人员姓名，且不能包含空格！");
        System.out.println("6：该程序会将Excel表中每个人的信息直接推送到槐荫人社企业号的<个人空间>模块！");
        System.out.println(" ");
        String result;
        do {
            // 输出提示文字
            System.out.print("请输入Excel文件名：");
            InputStreamReader is_reader = new InputStreamReader(System.in);
            result = new BufferedReader(is_reader).readLine();
        } while (result.equals("")); // 当用户输入无效的时候，反复提示要求用户输入
        File file = new File("C:\\" + result + ".xlsx");
        if (!file.exists()) {
            System.out.print("C:\\" + result + ".xlsx文件不存在！");
            System.out.print("按回车关闭程序...");
            while (true) {
                if (System.in.read() == '\n')
                    System.exit(0);
            }
        } else {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("c:\\" + result + ".xlsx"));
            XSSFSheet sheet = workbook.getSheetAt(0);
            int count = sheet.getRow(0).getPhysicalNumberOfCells();//列数
            int total = sheet.getLastRowNum();//行数
            String[] title = new String[count];
            String[] data = new String[count];
            String[] name = new String[total + 1];
            for (int i = 0; i < count; i++) {
                title[i] = sheet.getRow(0).getCell(i).toString();
            }
            for (int j = 1; j < total + 1; j++) {
                name[j] = sheet.getRow(j).getCell(0).getStringCellValue();
            }
            for (int i = 1; i < total + 1; i++) {
                for (int l = 0; l < count; l++) {
                    data[l] = sheet.getRow(i).getCell(l).toString();
                }
                StringBuffer messages = new StringBuffer();
                messages.append(result).append("\n");
                for (int j = 0; j < count; j++) {
                    messages.append(title[j] + "：" + data[j]).append("\n");
                }
                messages.append("如有异议请咨询相关工作人员！");
                String userId = new PinyinTool().toPinYin(name[i], "", PinyinTool.Type.FIRSTUPPER);
                try {
                    WxCpUser user = WxService.getMe().userGet(userId);
                    WxCpMessage message = WxCpMessage
                            .TEXT()
                            .agentId(ParamesAPI.AgentId) // 企业号应用ID
                            .toUser(userId)
                            .content(messages.toString())
                            .build();
                    WxService.getMe().messageSend(message);
                } catch (WxErrorException e) {
                    System.out.println(name[i] + "-发送失败(用户userId不存在)");
                }
            }
        }
        System.out.println("发送完成！");
        System.out.println("按回车键退出程序...");
        while (true) {
            if (System.in.read() == '\n')
                System.exit(0);
        }
    }
}

