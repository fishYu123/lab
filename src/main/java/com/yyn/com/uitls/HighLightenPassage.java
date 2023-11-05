package com.yyn.com.uitls;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class HighLightenPassage {

    private Map<String, List<String>> map;
    private List<String> entityList;



    /** 读取excel中的数据到list中
     *
     * @param fileName 需要处理的excel文件路径
     * @return 返回一个实体名字集合
     */
    public  void getEntityFromExcel(String fileName){
        try(    XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream(fileName));)
        {
            XSSFSheet sheet = xssfWorkbook.getSheet("WorkSheet");
            XSSFRow row ;
            XSSFCell cell;
            List<String> entityList = new ArrayList<>();
            Map<String, List<String>> map = new HashMap<>();
            List<String> list;
            for (int i=1; i<=sheet.getLastRowNum(); i++){
                row = sheet.getRow(i);
                list = new ArrayList<>();
                for(int j = 0; j<row.getPhysicalNumberOfCells(); j++){


                    cell = row.getCell(j);

                    String entity = cell.getStringCellValue();
                    list.add(entity);

                    if(!(entity.equals(""))){
                        entityList.add(entity);
                    }
                }
                map.put(row.getCell(0).getStringCellValue(), list);

            }
            this.map = map;
            this.entityList = entityList;
        }
        catch (Exception e){
            e.printStackTrace();
        }


    }


    /**将passage中与库里相同的实体名字进行高亮
     *
     * @param passage 接收的文章
     * @param dict 库里的实体名字集合
     * @return 返回处理好的高亮文章（添加html标签），直接放在浏览器中即可查看效果
     */
    public  String processPassage(String passage, List<String> dict){

        String pre = "<a href = \"";
        //实体的正式名称
        String obj;
        String end = "\" style=\"background-color: yellow\">";
        String last = "</a>";
        String processedPassage ="";

        for(String text : passage.split("，|。")){

            for(String entity : dict){

                if(text.contains(entity)){
                    obj = this.getSubject(entity, this.getMap());
                    Integer preIdx = text.indexOf(entity);
                    Integer endIdx = text.indexOf(entity) + entity.length();
                    String res = text.substring(0, preIdx) + pre + obj +end+  entity  + last+ text.substring(endIdx, text.length());
                    passage = passage.replace(text, res);
                }
            }
        }
        return passage;
    }

    /**
     *
     * @param entity 实体的名称或别名
     * @param map {实体正式名字，实体别名列表}
     * @return 实体正式名字
     */
    public  String getSubject (String entity, Map<String, List<String>> map){

        for(String key : map.keySet()){
            List<String> list = map.get(key);
            if(list.contains(entity)){
                return key;
            }
        }
        return "";
    }

    public void setHref(String filePath){
        try(XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream(filePath));
            FileOutputStream fos = new FileOutputStream(filePath))
        {
            XSSFSheet sheet = xssfWorkbook.getSheet("WorkSheet");
            XSSFRow row ;
            XSSFCell cell;
            List<String> entityList = new ArrayList<>();
            Map<String, List<String>> map = new HashMap<>();
            List<String> list;
            for (int i=1; i<=sheet.getLastRowNum(); i++){
                row = sheet.getRow(i);
                row.createCell(0).setCellValue("localhost:8080/home/list.html?entityType=&keyword="+row.getCell(1).getStringCellValue()+"&relation=&country=&entitySx=&sxMin=&sxMax=");


                for(int j = 0; j<row.getPhysicalNumberOfCells(); j++){


                    cell = row.getCell(j);
                    String entity = cell.getStringCellValue();
//
//                    if(!(entity.equals(""))){
//                        cell.setCellValue(entity+"(localhost:8080/home/list.html?entityType=&keyword="+this.getSubject(entity, this.getMap())+"&relation=&country=&entitySx=&sxMin=&sxMax=)");
//                    }
                }


            }
            xssfWorkbook.write(fos);
        }
        catch (Exception e){

        }
    }

    public List<String> getEntityList() {
        return entityList;
    }

    public void setEntityList(List<String> entityList) {
        this.entityList = entityList;
    }

    public Map<String, List<String>> getMap() {
        return map;
    }

    public void setMap(Map<String, List<String>> map) {
        this.map = map;
    }
}
