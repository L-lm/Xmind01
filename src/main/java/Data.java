import net.sf.json.JSONArray;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Data {
    String path;
    String fileName;
    //只做数据处理
    public List<Map<String, Object>> setPath(String path) throws IOException {
        this.path = path;
        File file = new File(path);
        this.fileName = file.getName();
        fileName = fileName.substring(0, fileName.indexOf("."));
        //System.out.println(this.fileName);
        List<Map<String, Object>> mapsList = null;
        String result = null;
        if (!file.exists()) {
            System.out.println("文件名不存在，请确认无误后重试！");
        } else {
            if (!file.canRead()) {
                System.out.println("该文件没有读取权限，请确认无误后重试！");
            } else {
                // 打开特定类型的输入流，读取文件的字节流 读取的内容in是一个地址：java.io.FileInputStream@19469ea2
                InputStream in = new FileInputStream(file);
                //System.out.println((char) in.read());
                //读取
                XSSFWorkbook sheets = new XSSFWorkbook(in);
                //System.out.println( sheets);
                //获取第一个sheet
                XSSFSheet sheet = sheets.getSheetAt(0);
                //getRow(0)读取第一行表头 A1/B1...
                XSSFRow titile = sheet.getRow(0);
                //System.out.println(sheet.getRow(0));
                System.out.println(sheet.getPhysicalNumberOfRows());
                //定义二维数组，存储读取出来的数据
                mapsList = new ArrayList<>();
                for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                    //从第二行开始
                    XSSFRow row = sheet.getRow(i);
                    //System.out.println(row.getPhysicalNumberOfCells());
                    //用于存储当前行的所有数据，以键-值的形式存储 HashMap允许存储相同的元素
                    Map<String, Object> map = new LinkedHashMap<>();
                    //循环遍历当前行
                    //row.getPhysicalNumberOfCells()只统计有内容总列数
                    //getLastCellNum 如果最后有数据的列为n，则返回n-1
                    for (int j = 0; j < row.getLastCellNum() + 1; j++) {
                        //获取这行数据中，对应列的键-值
                        XSSFCell title1 = titile.getCell(j);
                        //System.out.println(title1);
                        XSSFCell cell = row.getCell(j);
                        //System.out.println(cell);
                        //将获取到的值进行类型转换
                        if (cell == null) {
                            continue;
                        }
                        //修改单元格类型-文本类型
                        cell.setCellType(CellType.STRING);
                        //map里面只能是String Object类型的数据，现在的title1和cell的类型是XSSFCell类型 所以需要装换一下
                        String titleName = title1.getStringCellValue();
                        String valueName = cell.getStringCellValue();
                        //替换表格中的换行符
                        //if (valueName.contains("\n")) {
                         //   valueName = valueName.replace("\n", "***");
                        //}
                        map.put(titleName, valueName);
                    }
                    System.out.println(map);
                    mapsList.add(map);
                }
            }
        }
        return mapsList;
    }
}
