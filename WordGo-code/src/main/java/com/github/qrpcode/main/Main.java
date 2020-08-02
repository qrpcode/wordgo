package com.github.qrpcode.main;

import com.github.qrpcode.domain.WordGo;

/**
 * HelloWord程序
 * @author Qiruipeng
 */
public class Main {
    public static void main(String[] args) {
        Long time1 = System.currentTimeMillis();
        WordGo wordGo = new WordGo();
        wordGo.add("Hello World", "font-size: 20; color: #FF0000; line-height:150%;");
        wordGo.create("C:\\HelloWorld.docx");
        Long time2 = System.currentTimeMillis();
        System.out.println("C盘下已经有了HelloWord~ / 解锁更多技能 ↓");
        System.out.println("使用文档：https://github.com/qrpcode/wordgo");
        System.out.println("执行时间：" + (time2 - time1));
    }
}
