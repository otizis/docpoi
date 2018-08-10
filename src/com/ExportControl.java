package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Java Web根据模板导出word文件（spring MVC controller实现）
 *
 * @author lmb
 * @date 2017-3-14
 */
public class ExportControl
{

    public static void main(String[] args) throws FileNotFoundException
    {
        Map<String, Object> params = new HashMap<String, Object>();
        params.put("${name}", "safasdfa");
        params.put("${sex}", "111");

        List<Map<String, Object>> applyList = new ArrayList<Map<String, Object>>();
        for (int i = 0; i < 4; i++)
        {
            Map<String, Object> apply = new HashMap<String, Object>();
            apply.put("${name}", "safasdfa" + i);
            apply.put("${sex}", "111"+i);
            applyList.add(apply);
        }
        params.put("applyList", applyList);

        new ExportControl().exportWord(params);
    }

    /**
     * 导出word
     */
    public void exportWord(Map<String, Object> params) throws FileNotFoundException
    {


        XwpfUtil xwpfUtil = new XwpfUtil();
        //读入word模板
        InputStream is = new FileInputStream(new File("tpl.docx"));
        xwpfUtil.exportWord(params, is, xwpfUtil);
    }


}