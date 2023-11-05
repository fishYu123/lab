package com.yyn;

import com.yyn.com.uitls.HighLightenPassage;

public class myMain {
    public static void main(String[] args) {
        HighLightenPassage highLightenPassage = new HighLightenPassage();
        highLightenPassage.getEntityFromExcel("D:\\Code\\java\\lab\\src\\main\\java\\com\\yyn\\com\\装备数据别名绰号舷号.xlsx");
        highLightenPassage.setHref("D:\\Code\\java\\lab\\src\\main\\java\\com\\yyn\\com\\装备数据别名绰号舷号.xlsx");
        System.out.println("switch yyn branch");
        System.out.println("yyn 123");
        System.out.println("yyn 456");
    }
}
