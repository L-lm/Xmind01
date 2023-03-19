import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xmind.core.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws IOException, InvalidFormatException, CoreException {
        Data data= new Data();
        //excel存放的目录
        String path="/Users/may/Desktop/";
        List<Map<String, Object>> dataList = new ArrayList<>();
        //目录拼接文件名 获取到对应的文件
        dataList=data.setPath(path+"test_case01.xlsx");
        //创建一个空白界面
        IWorkbookBuilder workbookBuilder= (IWorkbookBuilder) Core.getWorkbookBuilder();
        //创建工作溥
        IWorkbook iWorkbook=workbookBuilder.createWorkbook(data.fileName);
        //获取默认sheet
        ISheet primarySheet=iWorkbook.getPrimarySheet();
        //创建一个流程图的主题
        ITopic rootTopic= primarySheet.getRootTopic();
        //设置成为正确的逻辑图
        rootTopic.setStructureClass("org.xmind.ui.logic.right");
        rootTopic.setTitleText(data.fileName);
        //忽略表头
        for (int i=1;i<dataList.size();i++){
            //因为第一列是表头 不需要显示在逻辑图中 所以忽略第一行的数据 i从0开始
            //先定义一个父节点的位置，父结点也就是当前数据需要链接的结点，就当前位置 父结点就是根结点
            ITopic lastTopic = rootTopic;
            for (String value : dataList.get(i).keySet()) {
                //标记是否存在相同结点
                boolean flag=false;
                String key=value;
                //就当前数据创建一个新结点
                ITopic topic = iWorkbook.createTopic();
                topic.setTitleText((String) dataList.get(i).get(key));
                //caseId不用显示在逻辑图中，所以对key==caseId 数据不做处理
                if (key.equals("caseId")) {
                    continue;
                }
                //创建一list对象 用来装父结点的所子结点
                List<ITopic> chil= (List<ITopic>) lastTopic.getAllChildren();
                //如果chil长度<1,证明当前父结点没有子结点，直接父结点中添加就行
                if (chil.size()<1){
                    //将当前结点链接到父结点上
                    lastTopic.add(topic,ITopic.ATTACHED);
                    //将当前结点当作下一个新结点的父结点
                    lastTopic = topic;
                }
                else {
                    //如果chil长度>1,可能存在当前结点topic和lasttopic结点相同，
                    for (int k=0;k<chil.size();k++){
                        if (topic.getTitleText()==chil.get(k).getTitleText()) {
                            //如果相同 就将相同结点的topic 赋给 lasttopic
                            //并且将flag=true 表示存在相同结点
                            lastTopic = chil.get(k);
                            flag = true;
                            break;
                        }
                    }
                    if (!flag){//如果flag==flase，证明不存在相同结点，所以即使循环结束也依旧要把当前结点 topic链接到lasttopic 并且更新lasttopic
                        lastTopic.add(topic,ITopic.ATTACHED);
                        lastTopic = topic;
                    }
                }
            }
        }
        File file1 = new File(path+data.fileName+".xmind");
        //判断该路径下是否存在该文件，如果存在那么删除，如果不存在 直接新建
        if (file1.exists()){
            file1.delete();
        }
        iWorkbook.save(path+data.fileName+".xmind");


    }
}
